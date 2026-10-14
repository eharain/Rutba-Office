/**
 * Ink as Office keeps it: an InkML part (`application/inkml+xml`, W3C
 * InkML) holding each stroke's points and its brush, colour, width, a
 * highlighter's rectangular tip and mask pen, which a drawing on a page
 * or a slide names through a content part (Word's `w14:contentPart`,
 * PowerPoint's `p:contentPart`).
 *
 * Written here as Word and PowerPoint write it (points in hundredths of a
 * millimetre, 1000 to the centimetre); read here as they write it too,
 * explicit values and the first and second differences their traces are
 * packed with, so a stroke drawn in Word or PowerPoint is drawn here.
 */

const HIMETRIC_PER_PX = 2540 / 96;

export const INK_CT = 'application/inkml+xml';
export const INK_REL = 'http://schemas.microsoft.com/office/2011/relationships/ink';
export const WPI_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk';
export const INK_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk';

const num = (v) => (Math.round(v * 1000) / 1000).toString();

/**
 * One stroke as an InkML part: `points` in px from the box's own top-left,
 * the pen's colour (#RRGGBB), width in points, and whether it is a
 * highlighter (a square tip that marks under the words, as Office's does).
 */
export function inkmlXml(strokes) {
  const brushes = [];
  const traces = [];
  strokes.forEach((s, i) => {
    const widthCm = Math.max(0.01, ((Number(s.width) || 1) / 72) * 2.54);
    const high = s.highlighter || (s.alpha != null && s.alpha < 1 && s.cap === 'sq');
    brushes.push(`<inkml:brush xml:id="br${i}"><inkml:brushProperty name="width" value="${num(widthCm)}" units="cm"/><inkml:brushProperty name="height" value="${num(widthCm)}" units="cm"/>`
      + `<inkml:brushProperty name="color" value="${String(s.color || '#000000').toUpperCase()}"/>`
      + (high ? '<inkml:brushProperty name="tip" value="rectangle"/><inkml:brushProperty name="rasterOp" value="maskPen"/>' : '')
      + (s.alpha != null && s.alpha < 1 && !high ? `<inkml:brushProperty name="transparency" value="${Math.round((1 - s.alpha) * 255)}"/>` : '')
      + '</inkml:brush>');
    const pts = (s.points || []).map(([x, y]) => `${Math.round(x * HIMETRIC_PER_PX)} ${Math.round(y * HIMETRIC_PER_PX)}`);
    traces.push(`<inkml:trace contextRef="#ctx0" brushRef="#br${i}">${pts.join(', ')}</inkml:trace>`);
  });
  return '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions><inkml:context xml:id="ctx0"><inkml:inkSource xml:id="inkSrc0"><inkml:traceFormat>'
    + '<inkml:channel name="X" type="integer" max="32767" units="cm"/><inkml:channel name="Y" type="integer" max="32767" units="cm"/></inkml:traceFormat>'
    + '<inkml:channelProperties><inkml:channelProperty channel="X" name="resolution" value="1000" units="1/cm"/><inkml:channelProperty channel="Y" name="resolution" value="1000" units="1/cm"/></inkml:channelProperties>'
    + '</inkml:inkSource></inkml:context>' + brushes.join('') + '</inkml:definitions>' + traces.join('') + '</inkml:ink>';
}

/** One trace's points, its packed values undone: explicit, first and second differences, per channel. */
export function readTrace(text, channels = 2) {
  const points = [];
  const prev = new Array(channels).fill(0);
  const vel = new Array(channels).fill(0);
  const mode = new Array(channels).fill('!');
  for (const raw of String(text).split(',')) {
    const tokens = [...raw.matchAll(/([!'"])?\s*([-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?)/g)];
    if (!tokens.length) continue;
    const point = [];
    tokens.slice(0, channels).forEach((t, c) => {
      if (t[1]) mode[c] = t[1];
      const v = Number(t[2]);
      let x;
      if (mode[c] === "'") { vel[c] = v; x = prev[c] + v; } else if (mode[c] === '"') { vel[c] += v; x = prev[c] + vel[c]; } else { vel[c] = v - prev[c]; x = v; }
      prev[c] = x;
      point.push(x);
    });
    if (point.length >= 2) points.push([point[0], point[1]]);
  }
  return points;
}

/**
 * An InkML part read: each stroke's points (in its own units), colour,
 * width in the same units, and whether it is a highlighter; and the box
 * round them all.
 */
export function readInkml(xml) {
  const source = String(xml || '');
  const brushes = new Map();
  for (const b of source.matchAll(/<inkml:brush\b[^>]*xml:id="([^"]*)"[^>]*>([\s\S]*?)<\/inkml:brush>/g)) {
    const prop = (name) => new RegExp('<inkml:brushProperty\\b[^>]*name="' + name + '"[^>]*value="([^"]*)"(?:[^>]*units="([^"]*)")?').exec(b[2]);
    const width = prop('width');
    const widthCm = width ? Number(width[1]) * (width[2] === 'mm' ? 0.1 : width[2] === 'in' ? 2.54 : 1) : 0.05;
    const transparency = Number(prop('transparency')?.[1] ?? 0);
    brushes.set(b[1], {
      color: prop('color')?.[1] || '#000000',
      widthCm,
      highlighter: /rectangle/.test(prop('tip')?.[1] || '') || /maskPen/i.test(prop('rasterOp')?.[1] || ''),
      alpha: transparency ? Math.max(0, 1 - transparency / 255) : 1,
    });
  }
  // Units per centimetre, from the X channel's resolution (Office's 1000/cm when it states none).
  const res = /<inkml:channelProperty\b[^>]*channel="X"[^>]*name="resolution"[^>]*value="([^"]*)"/.exec(source);
  const perCm = res ? Number(res[1]) || 1000 : 1000;
  const channels = Math.max(2, [...(/<inkml:traceFormat\b[^>]*>([\s\S]*?)<\/inkml:traceFormat>/.exec(source)?.[1] ?? '').matchAll(/<inkml:channel\b/g)].length);
  const strokes = [];
  for (const t of source.matchAll(/<inkml:trace\b([^>]*)>([\s\S]*?)<\/inkml:trace>/g)) {
    const ref = /brushRef="#?([^"]*)"/.exec(t[1])?.[1];
    const brush = brushes.get(ref) || { color: '#000000', widthCm: 0.05, highlighter: false, alpha: 1 };
    const points = readTrace(t[2], channels);
    if (points.length) strokes.push({ ...brush, width: brush.widthCm * perCm, points });
  }
  const xs = strokes.flatMap((s) => s.points.map((p) => p[0]));
  const ys = strokes.flatMap((s) => s.points.map((p) => p[1]));
  const bounds = xs.length ? { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) } : null;
  return { strokes, bounds, perCm };
}

/**
 * The strokes of an InkML part as a DrawingML freeform's path in a box
 * `cx` by `cy` EMU, the shape a reader that draws shapes draws, the
 * strokes laid on the box as their own bounds, a stroke's width kept in.
 * Answers `{ path, line }` for each stroke.
 */
export function inkPaths(xml, cx, cy) {
  const { strokes, bounds, perCm } = readInkml(xml);
  if (!strokes.length || !bounds) return [];
  const pad = Math.max(...strokes.map((s) => s.width)) / 2;
  const w = Math.max(1, bounds.w + pad * 2);
  const h = Math.max(1, bounds.h + pad * 2);
  const k = Math.min(cx / w, cy / h);
  const at = ([x, y]) => `<a:pt x="${Math.round((x - bounds.x + pad) * k)}" y="${Math.round((y - bounds.y + pad) * k)}"/>`;
  return strokes.map((s) => {
    const pts = s.points.length === 1 ? [s.points[0], [s.points[0][0] + 1, s.points[0][1]]] : s.points;
    let path = `<a:moveTo>${at(pts[0])}</a:moveTo>`;
    for (let i = 1; i < pts.length - 1; i++) path += `<a:quadBezTo>${at(pts[i])}${at([(pts[i][0] + pts[i + 1][0]) / 2, (pts[i][1] + pts[i + 1][1]) / 2])}</a:quadBezTo>`;
    path += `<a:lnTo>${at(pts[pts.length - 1])}</a:lnTo>`;
    // The pen's width in points: its centimetres, at the scale the strokes were laid on.
    const widthPt = Math.max(0.25, ((s.width * k) / 12700));
    return { path: `<a:path w="${Math.round(cx)}" h="${Math.round(cy)}" fill="none" extrusionOk="0">${path}</a:path>`, color: s.color.replace('#', '').toUpperCase(), widthPt, alpha: s.highlighter ? 0.5 : s.alpha, cap: s.highlighter ? 'sq' : 'rnd', perCm };
  });
}
