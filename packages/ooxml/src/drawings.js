/**
 * The drawings in a Word body — pictures, shapes, text boxes and groups —
 * found by the id Word gives each one (`wp:docPr id`), read for what the
 * page and the Arrange commands need, and rewritten in place: only the
 * element a command changes is touched, and everything else in the drawing
 * (effects, locks, the extensions a newer Word writes) rides through.
 *
 * A drawing is `<w:drawing>` holding `wp:inline` (a character in the line)
 * or `wp:anchor` (floating, with a position, a wrap and a place in the
 * z-order). Word 2010 and later write a shape or a text box inside
 * `mc:AlternateContent`, the DrawingML in `mc:Choice` and a VML copy in
 * `mc:Fallback` for older readers; the Choice is the drawing, and the
 * Fallback is written again from it whenever the Choice changes.
 *
 * Pure string work, no package: `document.js` owns the body and the parts.
 */
import { attrs, esc } from './package.js';

export const EMU_PER_PX = 9525;
const toPx = (emu) => Number(emu) / EMU_PER_PX;
const toEmu = (px) => Math.round(Number(px) * EMU_PER_PX);

/** The namespaces a floating drawing needs in scope, declared on the document as Word declares them. */
export const DRAWING_NS = {
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
};

/** Word's own ceiling for `relativeHeight` is a 32-bit value; its first drawing sits at 251658240. */
export const Z_BASE = 251658240;
export const Z_STEP = 1024;

/**
 * Every top-level drawing in a stretch of body XML, in document order —
 * not one in an `mc:Fallback` (the VML twin of a drawing already counted)
 * and not one inside a text box's content (a picture in a box's words is
 * part of those words, not a drawing on the page). Each is its `<w:drawing>`
 * span, the unit that holds it in its run (the AlternateContent when there
 * is one, else the drawing), and the run.
 */
export function findDrawings(body) {
  const out = [];
  const re = /<(\/?)(mc:Fallback|w:txbxContent|w:drawing)\b[^>]*?(\/?)>/g;
  let fallback = 0;
  let box = 0;
  let depth = 0;
  let start = -1;
  let m;
  while ((m = re.exec(body))) {
    const close = m[1] === '/';
    if (m[3] === '/') continue;
    const name = m[2];
    if (name === 'mc:Fallback') { fallback = Math.max(0, fallback + (close ? -1 : 1)); continue; }
    if (name === 'w:txbxContent') { box = Math.max(0, box + (close ? -1 : 1)); continue; }
    if (!close) {
      if (depth === 0 && fallback === 0 && box === 0) start = m.index;
      depth += 1;
    } else {
      depth = Math.max(0, depth - 1);
      if (depth === 0 && start >= 0) {
        out.push(locate(body, start, m.index + m[0].length));
        start = -1;
      }
    }
  }
  return out;
}

/** A drawing's span, its unit and its run, from where its `<w:drawing>` starts and ends. */
function locate(body, start, end) {
  const xml = body.slice(start, end);
  let unitStart = start;
  let unitEnd = end;
  const acOpen = body.lastIndexOf('<mc:AlternateContent', start);
  if (acOpen >= 0 && body.lastIndexOf('</mc:AlternateContent>', start) < acOpen) {
    const close = body.indexOf('</mc:AlternateContent>', end);
    if (close >= 0) { unitStart = acOpen; unitEnd = close + '</mc:AlternateContent>'.length; }
  }
  let runStart = -1;
  const re = /<w:r[\s>]/g;
  let m;
  const from = Math.max(0, unitStart - 4000);
  re.lastIndex = from;
  while ((m = re.exec(body)) && m.index < unitStart) runStart = m.index;
  if (runStart >= 0 && body.slice(runStart, unitStart).includes('</w:r>')) runStart = -1;
  const runClose = body.indexOf('</w:r>', unitEnd);
  const runEnd = runStart >= 0 && runClose >= 0 ? runClose + '</w:r>'.length : -1;
  const idMatch = /<wp:docPr\b[^>]*?\bid="(\d+)"/.exec(xml);
  return {
    id: idMatch ? Number(idMatch[1]) : null,
    start, end, xml,
    unitStart, unitEnd, alternate: unitStart !== start,
    runStart, runEnd,
  };
}

/** The drawing's own top transform — the first `a:xfrm` under its graphic. */
function topXfrm(xml) {
  const at = xml.indexOf('<a:graphicData');
  if (at < 0) return null;
  const re = /<a:xfrm\b([^>]*?)(\/?)>/g;
  re.lastIndex = at;
  const m = re.exec(xml);
  if (!m) return null;
  return { index: m.index, tag: m[0], attrs: attrs(m[1]) };
}

/**
 * What the page and the Arrange commands need of one drawing: its id and
 * name, what it is, whether it floats, where, how the words treat it, its
 * place in the z-order, its size and its turn.
 */
export function readDrawing(xml) {
  const head = /<wp:(anchor|inline)\b([^>]*)>/.exec(xml);
  const a = head ? attrs(head[2]) : {};
  const anchored = head?.[1] === 'anchor';
  const docPr = attrs(/<wp:docPr\b([^>]*?)\/?>/.exec(xml)?.[1] ?? '');
  const ext = attrs(/<wp:extent\b([^>]*)\/>/.exec(xml)?.[1] ?? '');
  const uri = /<a:graphicData\b[^>]*\buri="([^"]*)"/.exec(xml)?.[1] ?? '';
  const kind = /wordprocessingGroup/.test(uri) ? 'group'
    : /\/picture$/.test(uri) ? 'picture'
      : /wordprocessingShape/.test(uri) ? (/<wps:txbx\b/.test(xml) ? 'textbox' : 'shape')
        : /chart/.test(uri) ? 'chart'
          : /wordprocessingCanvas/.test(uri) ? 'canvas' : 'other';
  const xf = topXfrm(xml);
  const rot = xf?.attrs.rot ? Number(xf.attrs.rot) / 60000 : 0;
  return {
    id: docPr.id != null ? Number(docPr.id) : null,
    name: docPr.name ?? null,
    descr: docPr.descr ?? null,
    hidden: docPr.hidden === '1' || docPr.hidden === 'true',
    kind,
    anchored,
    widthPx: ext.cx ? toPx(ext.cx) : null,
    heightPx: ext.cy ? toPx(ext.cy) : null,
    relativeHeight: anchored && a.relativeHeight != null ? Number(a.relativeHeight) : null,
    behind: anchored && (a.behindDoc === '1' || a.behindDoc === 'true'),
    rot: Number.isFinite(rot) ? ((rot % 360) + 360) % 360 : 0,
    flipH: xf?.attrs.flipH === '1' || xf?.attrs.flipH === 'true',
    flipV: xf?.attrs.flipV === '1' || xf?.attrs.flipV === 'true',
  };
}

/* ── rewriting one part of a drawing ─────────────────────────────────────── */

/** An attribute of the drawing's `wp:anchor` set (or removed with null). */
export function withAnchorAttrs(xml, values) {
  return xml.replace(/<wp:anchor\b([^>]*)>/, (tag, inner) => {
    let out = inner;
    for (const [name, value] of Object.entries(values)) {
      const re = new RegExp('\\s' + name + '="[^"]*"');
      if (value == null) out = out.replace(re, '');
      else if (re.test(out)) out = out.replace(re, ' ' + name + '="' + esc(String(value)) + '"');
      else out += ' ' + name + '="' + esc(String(value)) + '"';
    }
    return '<wp:anchor' + out + '>';
  });
}

/** `wp:docPr`'s name or hidden flag, set in place. */
export function withDocPr(xml, { name, hidden } = {}) {
  return xml.replace(/<wp:docPr\b([^>]*?)(\/?)>/, (tag, inner, self) => {
    let out = inner;
    if (name !== undefined) {
      out = /\sname="/.test(out) ? out.replace(/\sname="[^"]*"/, ' name="' + esc(String(name)) + '"') : out + ' name="' + esc(String(name)) + '"';
    }
    if (hidden !== undefined) {
      out = out.replace(/\shidden="[^"]*"/, '');
      if (hidden) out += ' hidden="1"';
    }
    return '<wp:docPr' + out + self + '>';
  });
}

/** One axis of an anchor's position: relative to what, and an alignment or an offset in px. */
export function positionXml(axis, { rel, align = null, offsetPx = 0 }) {
  return '<wp:position' + axis + ' relativeFrom="' + rel + '">'
    + (align ? '<wp:align>' + align + '</wp:align>' : '<wp:posOffset>' + toEmu(offsetPx) + '</wp:posOffset>')
    + '</wp:position' + axis + '>';
}

export function withPosition(xml, axis, spec) {
  const re = new RegExp('<wp:position' + axis + '\\b[\\s\\S]*?<\\/wp:position' + axis + '>');
  return re.test(xml) ? xml.replace(re, positionXml(axis, spec)) : xml;
}

const WRAP_RE = /<wp:(wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom|wrapNone)\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/wp:\1>)/;

/** A wrap element as Word writes it; tight and through carry the drawing's box as their polygon. */
export function wrapXml(wrap, side = 'bothSides') {
  const polygon = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
  if (wrap === 'square') return '<wp:wrapSquare wrapText="' + side + '"/>';
  if (wrap === 'tight') return '<wp:wrapTight wrapText="' + side + '">' + polygon + '</wp:wrapTight>';
  if (wrap === 'through') return '<wp:wrapThrough wrapText="' + side + '">' + polygon + '</wp:wrapThrough>';
  if (wrap === 'topAndBottom') return '<wp:wrapTopAndBottom/>';
  return '<wp:wrapNone/>';
}

export function withWrap(xml, wrap, side) {
  const el = wrapXml(wrap, side);
  if (WRAP_RE.test(xml)) return xml.replace(WRAP_RE, el);
  // No wrap element yet: it goes after the effect extent, or the extent.
  if (/<wp:effectExtent\b[^>]*\/>/.test(xml)) return xml.replace(/(<wp:effectExtent\b[^>]*\/>)/, '$1' + el);
  return xml.replace(/(<wp:extent\b[^>]*\/>)/, '$1' + el);
}

/** The drawing's size: its extent and its own top transform's extent. */
export function withExtent(xml, widthPx, heightPx) {
  const cx = Math.max(1, toEmu(widthPx));
  const cy = Math.max(1, toEmu(heightPx));
  let out = xml.replace(/<wp:extent\b[^>]*\/>/, '<wp:extent cx="' + cx + '" cy="' + cy + '"/>');
  const xf = topXfrm(out);
  if (xf && !xf.tag.endsWith('/>')) {
    const close = out.indexOf('</a:xfrm>', xf.index);
    const inner = out.slice(xf.index + xf.tag.length, close);
    const next = /<a:ext\b[^>]*\/>/.test(inner)
      ? inner.replace(/<a:ext\b[^>]*\/>/, '<a:ext cx="' + cx + '" cy="' + cy + '"/>')
      : '<a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/>' + inner;
    out = out.slice(0, xf.index + xf.tag.length) + next + out.slice(close);
  }
  return out;
}

/** The drawing's turn and flips, on its own top transform (Word writes them there, not on the anchor). */
export function withTransform(xml, { rot, flipH, flipV }) {
  const xf = topXfrm(xml);
  if (!xf) return xml;
  let tagAttrs = xf.attrs;
  const next = { ...tagAttrs };
  if (rot !== undefined) {
    const deg = ((Math.round(Number(rot) || 0) % 360) + 360) % 360;
    if (deg) next.rot = String(deg * 60000); else delete next.rot;
  }
  if (flipH !== undefined) { if (flipH) next.flipH = '1'; else delete next.flipH; }
  if (flipV !== undefined) { if (flipV) next.flipV = '1'; else delete next.flipV; }
  const tag = '<a:xfrm' + Object.entries(next).map(([k, v]) => ' ' + k + '="' + esc(v) + '"').join('') + (xf.tag.endsWith('/>') ? '/>' : '>');
  return xml.slice(0, xf.index) + tag + xml.slice(xf.index + xf.tag.length);
}

/** The pieces a drawing is rebuilt from when it changes between inline and floating. */
function piecesOf(xml) {
  const pick = (re) => re.exec(xml)?.[0] ?? null;
  return {
    extent: pick(/<wp:extent\b[^>]*\/>/) ?? '<wp:extent cx="914400" cy="914400"/>',
    effectExtent: pick(/<wp:effectExtent\b[^>]*\/>/) ?? '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    docPr: pick(/<wp:docPr\b[^>]*\/>|<wp:docPr\b[^>]*>[\s\S]*?<\/wp:docPr>/) ?? '<wp:docPr id="1" name="Drawing"/>',
    frame: pick(/<wp:cNvGraphicFramePr\b[^>]*\/>|<wp:cNvGraphicFramePr\b[^>]*>[\s\S]*?<\/wp:cNvGraphicFramePr>/) ?? '<wp:cNvGraphicFramePr/>',
    graphic: pick(/<a:graphic\b[\s\S]*<\/a:graphic>/),
    open: /<w:drawing\b[^>]*>/.exec(xml)?.[0] ?? '<w:drawing>',
  };
}

const WP_DECL = ' xmlns:wp="' + DRAWING_NS.wp + '"';

/**
 * A floating drawing from its pieces: Word's anchor, in the schema's order —
 * simplePos, the two positions, the extent, the effect extent, the wrap, then
 * the drawing itself.
 */
export function anchorXml(pieces, {
  wrap = 'square', side = 'bothSides', behind = false, relativeHeight = Z_BASE,
  h = { rel: 'column', offsetPx: 0 }, v = { rel: 'paragraph', offsetPx: 0 },
  dist = { t: 0, b: 0, l: 114300, r: 114300 }, extras = '',
}) {
  return pieces.open + '<wp:anchor distT="' + dist.t + '" distB="' + dist.b + '" distL="' + dist.l + '" distR="' + dist.r + '" simplePos="0" relativeHeight="' + relativeHeight + '" behindDoc="' + (behind ? 1 : 0) + '" locked="0" layoutInCell="1" allowOverlap="1"' + (pieces.open.includes('xmlns:wp') ? '' : WP_DECL) + '>'
    + '<wp:simplePos x="0" y="0"/>'
    + positionXml('H', h) + positionXml('V', v)
    + pieces.extent + pieces.effectExtent + wrapXml(wrap, side) + pieces.docPr + pieces.frame + pieces.graphic
    + extras + '</wp:anchor></w:drawing>';
}

/** An inline drawing from a floating one — a character in the line again. */
export function toInline(xml) {
  const p = piecesOf(xml);
  if (!p.graphic) throw new Error('the drawing has no graphic');
  return p.open + '<wp:inline distT="0" distB="0" distL="0" distR="0"' + (p.open.includes('xmlns:wp') ? '' : WP_DECL) + '>' + p.extent + p.effectExtent + p.docPr + p.frame + p.graphic + '</wp:inline></w:drawing>';
}

/** A floating drawing from an inline one, at the position and wrap given. */
export function toAnchor(xml, spec) {
  const p = piecesOf(xml);
  if (!p.graphic) throw new Error('the drawing has no graphic');
  return anchorXml(p, spec);
}

/* ── a shape's look: fill, outline, and the text body's frame ─────────────── */

const FILL_RE = /<a:(?:solidFill|noFill|gradFill|blipFill|pattFill|grpFill)\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:(?:solidFill|gradFill|blipFill|pattFill)>)/;

/** The top shape's `wps:spPr`, as its span in the drawing. */
function spPrOf(xml) {
  const open = /<wps:spPr\b[^>]*>/.exec(xml);
  if (!open) return null;
  const close = xml.indexOf('</wps:spPr>', open.index);
  if (close < 0) return null;
  return { start: open.index, innerStart: open.index + open[0].length, innerEnd: close, end: close + '</wps:spPr>'.length };
}

const hex6 = (v) => {
  const h = String(v).replace(/^#/, '').toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(h)) throw new Error('a colour must be #RRGGBB: ' + v);
  return h;
};

/** The shape's fill — a colour, or none — before its outline, as the schema orders them. */
export function withShapeFill(xml, fill) {
  const sp = spPrOf(xml);
  if (!sp) return xml;
  let inner = xml.slice(sp.innerStart, sp.innerEnd);
  const ln = /<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/.exec(inner);
  let head = ln ? inner.slice(0, ln.index) : inner;
  const tail = ln ? inner.slice(ln.index) : '';
  head = head.replace(FILL_RE, '');
  const el = fill == null ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + hex6(fill) + '"/></a:solidFill>';
  const geom = /<a:(?:prstGeom|custGeom)\b[\s\S]*?<\/a:(?:prstGeom|custGeom)>|<a:(?:prstGeom|custGeom)\b[^>]*\/>/.exec(head);
  head = geom ? head.slice(0, geom.index + geom[0].length) + el + head.slice(geom.index + geom[0].length) : head + el;
  inner = head + tail;
  return xml.slice(0, sp.innerStart) + inner + xml.slice(sp.innerEnd);
}

/** The shape's outline: a colour and a weight in px, or none. */
export function withShapeLine(xml, { colour = null, widthPx = 1, none = false } = {}) {
  const sp = spPrOf(xml);
  if (!sp) return xml;
  let inner = xml.slice(sp.innerStart, sp.innerEnd);
  const w = Math.max(1, Math.round(Number(widthPx) * EMU_PER_PX));
  const el = none || colour == null
    ? '<a:ln><a:noFill/></a:ln>'
    : '<a:ln w="' + w + '"><a:solidFill><a:srgbClr val="' + hex6(colour) + '"/></a:solidFill></a:ln>';
  const ln = /<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/.exec(inner);
  if (ln) inner = inner.slice(0, ln.index) + el + inner.slice(ln.index + ln[0].length);
  else {
    // After the fill (or the geometry), before any effects.
    const fx = /<a:(?:effectLst|effectDag|scene3d|sp3d|extLst)\b/.exec(inner);
    inner = fx ? inner.slice(0, fx.index) + el + inner.slice(fx.index) : inner + el;
  }
  return xml.slice(0, sp.innerStart) + inner + xml.slice(sp.innerEnd);
}

/** The shape's fill and outline as written: hex strings, or null for none; `inherit` when the file names a style colour. */
export function readShapeLook(xml) {
  const sp = spPrOf(xml);
  const inner = sp ? xml.slice(sp.innerStart, sp.innerEnd) : '';
  const ln = /<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>|<a:ln\b([^>]*)\/>/.exec(inner);
  const lnAttrs = attrs(ln ? (ln[1] ?? ln[3] ?? '') : '');
  return {
    spPr: inner,
    lineXml: ln ? ln[0] : null,
    lineWidthPx: lnAttrs.w ? Number(lnAttrs.w) / EMU_PER_PX : (ln ? 0.75 * (96 / 72) : null),
  };
}

const INSET_DEFAULT = { l: 91440, t: 45720, r: 91440, b: 45720 };

/** A text box's frame for its words: the four insets (px), where the words sit up and down, and their direction. */
export function readBodyPr(xml) {
  const m = /<wps:bodyPr\b([^>]*?)(\/?)>/.exec(xml);
  const a = attrs(m?.[1] ?? '');
  const body = m && m[2] !== '/' ? xml.slice(m.index, xml.indexOf('</wps:bodyPr>', m.index)) : '';
  const ins = (k, key) => toPx(a[key] != null ? a[key] : INSET_DEFAULT[k]);
  return {
    insets: { l: ins('l', 'lIns'), t: ins('t', 'tIns'), r: ins('r', 'rIns'), b: ins('b', 'bIns') },
    anchor: a.anchor === 'ctr' ? 'middle' : a.anchor === 'b' ? 'bottom' : 'top',
    vert: a.vert === 'vert' || a.vert === 'eaVert' || a.vert === 'wordArtVert' ? 'vert' : a.vert === 'vert270' ? 'vert270' : 'horz',
    autoFit: /<a:spAutoFit\b/.test(body),
  };
}

/** A text box's body frame rewritten: insets in px, `anchor` top/middle/bottom, `vert` horz/vert/vert270, `autoFit`. */
export function withBodyPr(xml, { insets, anchor, vert, autoFit } = {}) {
  const m = /<wps:bodyPr\b([^>]*?)(\/?)>/.exec(xml);
  if (!m) return xml;
  const a = attrs(m[1]);
  if (insets) {
    for (const [k, key] of [['l', 'lIns'], ['t', 'tIns'], ['r', 'rIns'], ['b', 'bIns']]) {
      if (insets[k] != null) a[key] = String(Math.max(0, toEmu(insets[k])));
    }
  }
  if (anchor) a.anchor = anchor === 'middle' ? 'ctr' : anchor === 'bottom' ? 'b' : 't';
  if (vert) a.vert = vert === 'vert' ? 'vert' : vert === 'vert270' ? 'vert270' : 'horz';
  const open = '<wps:bodyPr' + Object.entries(a).map(([k, v]) => ' ' + k + '="' + esc(v) + '"').join('');
  let children = '';
  let end = m.index + m[0].length;
  if (m[2] !== '/') {
    const close = xml.indexOf('</wps:bodyPr>', m.index);
    children = xml.slice(m.index + m[0].length, close);
    end = close + '</wps:bodyPr>'.length;
  }
  if (autoFit !== undefined) {
    children = children.replace(/<a:(?:spAutoFit|noAutofit|normAutofit)\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:normAutofit>)/, '');
    children = (autoFit ? '<a:spAutoFit/>' : '<a:noAutofit/>') + children;
  }
  const el = children ? open + '>' + children + '</wps:bodyPr>' : open + '/>';
  return xml.slice(0, m.index) + el + xml.slice(end);
}

/* ── text boxes, as Word 2010 and later write them ───────────────────────── */

const pt = (px) => Math.round(Number(px) * 0.75 * 100) / 100;

/**
 * The VML copy of a text box an older Word reads: a rectangle with the same
 * size, place, fill, outline and words. Written from the DrawingML every time
 * the box changes, so the two never disagree.
 */
export function vmlTextBox({ id, name, widthPx, heightPx, h, v, wrap, behind, relativeHeight, fill, line, lineWidthPx, insets, anchor, content }) {
  const hRel = { page: 'page', margin: 'margin', column: 'text', character: 'char' }[h?.rel] || 'text';
  const vRel = { page: 'page', margin: 'margin', paragraph: 'text', line: 'line' }[v?.rel] || 'text';
  const z = behind ? -Math.max(1, (relativeHeight || Z_BASE) - Z_BASE + 1) : Math.max(1, (relativeHeight || Z_BASE) - Z_BASE + 1);
  const style = [
    'position:absolute',
    h?.align ? 'margin-left:0' : 'margin-left:' + pt(h?.offsetPx || 0) + 'pt',
    v?.align ? 'margin-top:0' : 'margin-top:' + pt(v?.offsetPx || 0) + 'pt',
    'width:' + pt(widthPx) + 'pt', 'height:' + pt(heightPx) + 'pt',
    'z-index:' + z, 'visibility:visible', 'mso-wrap-style:square',
    'mso-position-horizontal:' + (h?.align || 'absolute'), 'mso-position-horizontal-relative:' + hRel,
    'mso-position-vertical:' + (v?.align || 'absolute'), 'mso-position-vertical-relative:' + vRel,
    'v-text-anchor:' + (anchor === 'middle' ? 'middle' : anchor === 'bottom' ? 'bottom' : 'top'),
  ].join(';');
  const ins = insets || { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };
  const fillAttr = fill ? ' fillcolor="#' + hex6(fill) + '"' : ' filled="f"';
  const strokeAttr = line ? ' strokecolor="#' + hex6(line) + '" strokeweight="' + pt(lineWidthPx || 1) + 'pt"' : ' stroked="f"';
  const wrapType = wrap === 'topAndBottom' ? 'topAndBottom' : wrap === 'none' ? 'none' : wrap === 'tight' ? 'tight' : wrap === 'through' ? 'through' : 'square';
  return '<w:pict><v:rect id="' + esc(name || 'Text Box ' + id) + '" o:spid="_x0000_s' + (1024 + Number(id || 0)) + '" style="' + style + '"' + fillAttr + strokeAttr + '>'
    + '<v:textbox inset="' + [ins.l, ins.t, ins.r, ins.b].map((x) => pt(x) + 'pt').join(',') + '"><w:txbxContent>' + content + '</w:txbxContent></v:textbox>'
    + (wrapType === 'none' ? '' : '<w10:wrap type="' + wrapType + '"/>')
    + '</v:rect></w:pict>';
}

/**
 * A floating text box, whole: the run, the AlternateContent, Word's wps shape
 * in its anchor, and the VML twin. `paragraphs` is the box's content — one
 * or more `<w:p>`.
 */
export function textBoxRun(spec) {
  const {
    id, name = 'Text Box ' + id, widthPx, heightPx,
    h = { rel: 'column', offsetPx: 0 }, v = { rel: 'paragraph', offsetPx: 0 },
    wrap = 'square', behind = false, relativeHeight = Z_BASE,
    fill = 'FFFFFF', line = '000000', lineWidthPx = 1,
    insets = { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }, anchor = 'top', vert = 'horz', autoFit = false,
    paragraphs = '<w:p/>',
  } = spec;
  const cx = Math.max(1, toEmu(widthPx));
  const cy = Math.max(1, toEmu(heightPx));
  const fillXml = fill == null ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + hex6(fill) + '"/></a:solidFill>';
  const lineXml = line == null ? '<a:ln><a:noFill/></a:ln>' : '<a:ln w="' + Math.max(1, toEmu(lineWidthPx)) + '"><a:solidFill><a:srgbClr val="' + hex6(line) + '"/></a:solidFill></a:ln>';
  const bodyPr = '<wps:bodyPr rot="0" vert="' + (vert === 'vert' ? 'vert' : vert === 'vert270' ? 'vert270' : 'horz') + '" wrap="square"'
    + ' lIns="' + toEmu(insets.l) + '" tIns="' + toEmu(insets.t) + '" rIns="' + toEmu(insets.r) + '" bIns="' + toEmu(insets.b) + '"'
    + ' anchor="' + (anchor === 'middle' ? 'ctr' : anchor === 'bottom' ? 'b' : 't') + '" anchorCtr="0">'
    + (autoFit ? '<a:spAutoFit/>' : '<a:noAutofit/>') + '</wps:bodyPr>';
  const graphic = '<a:graphic xmlns:a="' + DRAWING_NS.a + '"><a:graphicData uri="' + DRAWING_NS.wps + '">'
    + '<wps:wsp><wps:cNvSpPr txBox="1"/>'
    + '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' + fillXml + lineXml + '</wps:spPr>'
    + '<wps:txbx><w:txbxContent>' + paragraphs + '</w:txbxContent></wps:txbx>'
    + bodyPr + '</wps:wsp></a:graphicData></a:graphic>';
  const drawing = anchorXml({
    open: '<w:drawing>',
    extent: '<wp:extent cx="' + cx + '" cy="' + cy + '"/>',
    effectExtent: '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    docPr: '<wp:docPr id="' + id + '" name="' + esc(name) + '"/>',
    frame: '<wp:cNvGraphicFramePr/>',
    graphic,
  }, { wrap, behind, relativeHeight, h, v, dist: { t: 45720, b: 45720, l: 114300, r: 114300 } });
  const fallback = vmlTextBox({ id, name, widthPx, heightPx, h, v, wrap, behind, relativeHeight, fill, line, lineWidthPx, insets, anchor, content: paragraphs });
  return '<w:r><mc:AlternateContent><mc:Choice Requires="wps">' + drawing + '</mc:Choice><mc:Fallback>' + fallback + '</mc:Fallback></mc:AlternateContent></w:r>';
}

/**
 * A text box's VML twin written again from its DrawingML — what an edit to
 * the box calls, so an older reader sees the same box with the same words.
 */
export function fallbackFor(choiceXml, colours = null) {
  const d = readDrawing(choiceXml);
  const layout = /<wp:positionH\b[^>]*relativeFrom="([^"]*)"[^>]*>([\s\S]*?)<\/wp:positionH>/.exec(choiceXml);
  const layoutV = /<wp:positionV\b[^>]*relativeFrom="([^"]*)"[^>]*>([\s\S]*?)<\/wp:positionV>/.exec(choiceXml);
  const axis = (m) => {
    if (!m) return { rel: 'column', offsetPx: 0 };
    const align = /<wp:align>([^<]*)<\/wp:align>/.exec(m[2])?.[1] ?? null;
    const off = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(m[2]);
    return { rel: m[1], align, offsetPx: off ? toPx(off[1]) : 0 };
  };
  const wrapEl = /<wp:(wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom|wrapNone)\b/.exec(choiceXml)?.[1];
  const wrap = { wrapSquare: 'square', wrapTight: 'tight', wrapThrough: 'through', wrapTopAndBottom: 'topAndBottom', wrapNone: 'none' }[wrapEl] || 'square';
  const look = readShapeLook(choiceXml);
  const solid = (frag) => {
    const m = /<a:solidFill>\s*<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(frag || '');
    if (m) return m[1];
    const s = /<a:solidFill>\s*<a:schemeClr\b[^>]*\bval="([^"]*)"/.exec(frag || '');
    return s && colours ? (colours[s[1]] ?? null) : null;
  };
  const fillFrag = look.spPr.replace(/<a:ln\b[\s\S]*?<\/a:ln>/, '');
  const body = readBodyPr(choiceXml);
  const content = /<w:txbxContent\b[^>]*>([\s\S]*)<\/w:txbxContent>/.exec(choiceXml)?.[1] ?? '<w:p/>';
  return vmlTextBox({
    id: d.id, name: d.name, widthPx: d.widthPx || 96, heightPx: d.heightPx || 48,
    h: axis(layout), v: axis(layoutV), wrap, behind: d.behind, relativeHeight: d.relativeHeight,
    fill: solid(fillFrag), line: /<a:noFill\b/.test(look.lineXml || '') ? null : solid(look.lineXml), lineWidthPx: look.lineWidthPx,
    insets: body.insets, anchor: body.anchor, content,
  });
}

/* ── groups: wpg:wgp ─────────────────────────────────────────────────────── */

/**
 * The shape a drawing becomes inside a group — its `pic:pic` or `wps:wsp`,
 * placed at `off` and sized `ext` (EMU, in the group's own space), turned as
 * it was. A group member has no `wp:docPr` of its own, so a shape carries
 * its id and name in `wps:cNvPr`, as Word writes it; a picture already has
 * them in `pic:cNvPr`.
 */
export function memberXml(drawingXml, { x, y, cx, cy }) {
  const d = readDrawing(drawingXml);
  let el = /<pic:pic\b[\s\S]*<\/pic:pic>/.exec(drawingXml)?.[0]
    ?? /<wps:wsp\b[\s\S]*<\/wps:wsp>/.exec(drawingXml)?.[0]
    ?? /<wpg:wgp\b[\s\S]*<\/wpg:wgp>/.exec(drawingXml)?.[0];
  if (!el) throw new Error('drawing ' + d.id + ' cannot go into a group');
  if (el.startsWith('<wpg:wgp')) el = '<wpg:grpSp' + el.slice('<wpg:wgp'.length, -'</wpg:wgp>'.length) + '</wpg:grpSp>';
  if (el.startsWith('<wps:wsp') && !/<wps:cNvPr\b/.test(el)) {
    el = el.replace(/^<wps:wsp\b[^>]*>/, (open) => open + '<wps:cNvPr id="' + d.id + '" name="' + esc(d.name || 'Shape ' + d.id) + '"/>');
  }
  const xf = /<a:xfrm\b([^>]*?)(\/?)>/.exec(el);
  const place = '<a:off x="' + Math.round(x) + '" y="' + Math.round(y) + '"/><a:ext cx="' + Math.max(1, Math.round(cx)) + '" cy="' + Math.max(1, Math.round(cy)) + '"/>';
  if (xf && xf[2] !== '/') {
    const close = el.indexOf('</a:xfrm>', xf.index);
    const inner = el.slice(xf.index + xf[0].length, close).replace(/<a:off\b[^>]*\/>/, '').replace(/<a:ext\b[^>]*\/>/, '');
    el = el.slice(0, xf.index + xf[0].length) + place + inner + el.slice(close);
  } else if (xf) {
    el = el.slice(0, xf.index) + '<a:xfrm' + xf[1] + '>' + place + '</a:xfrm>' + el.slice(xf.index + xf[0].length);
  } else {
    // A member with no transform of its own gets one, first in its shape properties.
    el = el.replace(/(<(?:pic|wps|wpg):(?:spPr|grpSpPr)\b[^>]*>)/, '$1<a:xfrm>' + place + '</a:xfrm>');
  }
  return el;
}

/** A group's members, each with its place in the group's space (EMU) and its XML. */
export function groupMembers(groupXml) {
  const wgp = /<wpg:wgp\b[^>]*>([\s\S]*)<\/wpg:wgp>/.exec(groupXml);
  if (!wgp) return { off: { x: 0, y: 0 }, ext: { cx: 1, cy: 1 }, chOff: { x: 0, y: 0 }, chExt: { cx: 1, cy: 1 }, members: [] };
  const inner = wgp[1];
  const grp = /<wpg:grpSpPr\b[^>]*>([\s\S]*?)<\/wpg:grpSpPr>/.exec(inner)?.[1] ?? '';
  const num = (re, s) => { const m = re.exec(s); return m ? attrs(m[1]) : {}; };
  const off = num(/<a:off\b([^>]*)\/>/, grp);
  const ext = num(/<a:ext\b([^>]*)\/>/, grp);
  const chOff = num(/<a:chOff\b([^>]*)\/>/, grp);
  const chExt = num(/<a:chExt\b([^>]*)\/>/, grp);
  // Top-level children only, depth-aware.
  const members = [];
  const re = /<(\/?)(pic:pic|wps:wsp|wpg:grpSp|wpc:wpc|w14:contentPart)\b[^>]*?(\/?)>/g;
  let depth = 0;
  let start = -1;
  let tag = null;
  let m;
  while ((m = re.exec(inner))) {
    if (m[3] === '/') continue;
    if (m[1] !== '/') {
      if (depth === 0) { start = m.index; tag = m[2]; }
      if (m[2] === tag) depth += 1;
    } else if (m[2] === tag) {
      depth -= 1;
      if (depth === 0) {
        const xml = inner.slice(start, m.index + m[0].length);
        const xf = /<a:xfrm\b([^>]*?)(\/?)>([\s\S]*?)(?:<\/a:xfrm>|$)/.exec(xml);
        const xa = attrs(xf?.[1] ?? '');
        const mo = num(/<a:off\b([^>]*)\/>/, xf?.[3] ?? '');
        const me = num(/<a:ext\b([^>]*)\/>/, xf?.[3] ?? '');
        const cNvPr = attrs(/<(?:pic|wps|wpg):cNvPr\b([^>]*?)\/?>/.exec(xml)?.[1] ?? '');
        members.push({
          tag, xml,
          id: cNvPr.id != null ? Number(cNvPr.id) : null,
          name: cNvPr.name ?? null,
          x: Number(mo.x || 0), y: Number(mo.y || 0), cx: Number(me.cx || 0), cy: Number(me.cy || 0),
          rot: xa.rot ? Number(xa.rot) / 60000 : 0,
          flipH: xa.flipH === '1', flipV: xa.flipV === '1',
          kind: tag === 'pic:pic' ? 'picture' : tag === 'wps:wsp' ? (/<wps:txbx\b/.test(xml) ? 'textbox' : 'shape') : 'group',
        });
        tag = null;
      }
    }
  }
  const n = (v, d) => (v == null || v === '' ? d : Number(v));
  return {
    off: { x: n(off.x, 0), y: n(off.y, 0) }, ext: { cx: n(ext.cx, 1), cy: n(ext.cy, 1) },
    chOff: { x: n(chOff.x, 0), y: n(chOff.y, 0) }, chExt: { cx: n(chExt.cx, n(ext.cx, 1)), cy: n(chExt.cy, n(ext.cy, 1)) },
    members,
  };
}

/** A group's graphic around its members, the group's space the members' box. */
export function groupGraphic(members, { cx, cy }) {
  return '<a:graphic xmlns:a="' + DRAWING_NS.a + '"><a:graphicData uri="' + DRAWING_NS.wpg + '">'
    + '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/><a:chOff x="0" y="0"/><a:chExt cx="' + cx + '" cy="' + cy + '"/></a:xfrm></wpg:grpSpPr>'
    + members.join('')
    + '</wpg:wgp></a:graphicData></a:graphic>';
}

/**
 * A member of a group as a drawing of its own again, floating where it sat
 * in the group: a picture keeps its `pic:pic`, a shape its `wps:wsp` (less
 * the group-only `wps:cNvPr`), a nested group becomes a group.
 */
export function memberToDrawing(member, { id, widthPx, heightPx, h, v, wrap, behind, relativeHeight, side }) {
  let el = member.xml;
  let uri = DRAWING_NS.pic.replace(/picture$/, 'picture');
  if (member.tag === 'pic:pic') uri = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  else if (member.tag === 'wps:wsp') {
    uri = DRAWING_NS.wps;
    el = el.replace(/<wps:cNvPr\b[^>]*\/>/, '');
  } else if (member.tag === 'wpg:grpSp') {
    uri = DRAWING_NS.wpg;
    el = '<wpg:wgp' + el.slice('<wpg:grpSp'.length, -'</wpg:grpSp>'.length) + '</wpg:wgp>';
    el = el.replace(/<wpg:cNvPr\b[^>]*\/>/, '');
  }
  const cx = Math.max(1, toEmu(widthPx));
  const cy = Math.max(1, toEmu(heightPx));
  // The member's transform now sits at its own origin.
  el = el.replace(/<a:off\b[^>]*\/>/, '<a:off x="0" y="0"/>').replace(/<a:ext\b[^>]*\/>/, '<a:ext cx="' + cx + '" cy="' + cy + '"/>');
  const name = member.name || (member.kind === 'picture' ? 'Picture ' : 'Shape ') + id;
  const graphic = '<a:graphic xmlns:a="' + DRAWING_NS.a + '"><a:graphicData uri="' + uri + '">' + el + '</a:graphicData></a:graphic>';
  return anchorXml({
    open: '<w:drawing>',
    extent: '<wp:extent cx="' + cx + '" cy="' + cy + '"/>',
    effectExtent: '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    docPr: '<wp:docPr id="' + id + '" name="' + esc(name) + '"/>',
    frame: '<wp:cNvGraphicFramePr/>',
    graphic,
  }, { wrap, side, behind, relativeHeight, h, v });
}

export { toPx as emuToPx, toEmu as pxToEmu };
