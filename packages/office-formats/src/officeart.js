// Office Art — the drawing records Word, Excel and PowerPoint share: a
// record's header and children, a shape's properties (OfficeArtFOPT), its
// colours, the numbered shape types as DrawingML presets, and a freeform's
// own outline, and a shape's fill, outline and shadow. Used by the
// PowerPoint reader, the Word reader's floating drawings and the Excel
// reader's drawings alike. Pure.

const u16 = (b, at) => (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
const i16 = (b, at) => { const v = u16(b, at); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, at) => (u16(b, at) | (u16(b, at + 2) << 16)) >>> 0;
const i32 = (b, at) => u32(b, at) | 0;
/** A record's header: version, instance, type, and where its body is. */
export function header(b, at) {
  const vi = u16(b, at);
  const len = u32(b, at + 4);
  return { ver: vi & 0xf, inst: vi >> 4, type: u16(b, at + 2), len, body: at + 8, end: Math.min(b.length, at + 8 + len) };
}
/** A container's children. */
export function children(b, h) {
  const out = [];
  if (!h) return out;
  let p = h.body;
  while (p + 8 <= h.end) {
    const c = header(b, p);
    out.push(c);
    if (c.len === 0 && c.ver !== 0xf) { p = c.body; continue; }
    p = c.body + c.len;
  }
  return out;
}
export const child = (b, h, type, inst = null) => children(b, h).find((c) => c.type === type && (inst == null || c.inst === inst)) || null;

/** An Office Art colour: its RGB, or a scheme entry when it says so. */
export function artColour(value, scheme) {
  const flags = value >>> 24;
  if (flags & 0x08) return scheme[value & 0xff] ?? null;
  if (flags & 0x10) return null; // a system colour this has no table for
  return [0, 8, 16].map((s) => ((value >>> s) & 0xff).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/* ── shapes ───────────────────────────────────────────────────────────── */

/** An OfficeArtFOPT's properties, by id: { op, complex? }. */
export function readFopt(doc, h) {
  const props = new Map();
  if (!h) return props;
  const n = h.inst;
  let complexAt = h.body + n * 6;
  for (let k = 0; k < n; k++) {
    const id = u16(doc, h.body + k * 6);
    const op = u32(doc, h.body + k * 6 + 2);
    const entry = { op };
    if (id & 0x8000) { entry.complex = doc.subarray(complexAt, Math.min(h.end, complexAt + op)); complexAt += op; }
    props.set(id & 0x3fff, entry);
  }
  return props;
}

// Office Art's shape types, by number, as DrawingML's preset geometries (LibreOffice's
// table of them, filter/source/msfilter/util.cxx). Lines and connectors, pictures,
// text boxes and WordArt are read as what they are, not as presets.
export const PRESETS = {
  1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon',
  10: 'octagon', 11: 'plus', 12: 'star5', 13: 'rightArrow', 14: 'rightArrow', 15: 'homePlate', 16: 'cube', 17: 'wedgeRoundRectCallout', 18: 'star16',
  19: 'arc', 21: 'plaque', 22: 'can', 23: 'donut', 41: 'callout1', 42: 'callout2', 43: 'callout3', 44: 'accentCallout1', 45: 'accentCallout2',
  46: 'accentCallout3', 47: 'borderCallout1', 48: 'borderCallout2', 49: 'borderCallout3', 50: 'accentBorderCallout1', 51: 'accentBorderCallout2',
  52: 'accentBorderCallout3', 53: 'ribbon', 54: 'ribbon2', 55: 'chevron', 56: 'pentagon', 57: 'noSmoking', 58: 'star8', 59: 'star16', 60: 'star32',
  61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout', 63: 'wedgeEllipseCallout', 64: 'wave', 65: 'foldedCorner', 66: 'leftArrow', 67: 'downArrow',
  68: 'upArrow', 69: 'leftRightArrow', 70: 'upDownArrow', 71: 'irregularSeal1', 72: 'irregularSeal2', 73: 'lightningBolt', 74: 'heart',
  76: 'quadArrow', 77: 'leftArrowCallout', 78: 'rightArrowCallout', 79: 'upArrowCallout', 80: 'downArrowCallout', 81: 'leftRightArrowCallout',
  82: 'upDownArrowCallout', 83: 'quadArrowCallout', 84: 'bevel', 85: 'leftBracket', 86: 'rightBracket', 87: 'leftBrace', 88: 'rightBrace',
  89: 'leftUpArrow', 90: 'bentUpArrow', 91: 'bentArrow', 92: 'star24', 93: 'stripedRightArrow', 94: 'notchedRightArrow', 95: 'blockArc',
  96: 'smileyFace', 97: 'verticalScroll', 98: 'horizontalScroll', 99: 'circularArrow', 100: 'notchedCircularArrow', 101: 'uturnArrow',
  102: 'curvedRightArrow', 103: 'curvedLeftArrow', 104: 'curvedUpArrow', 105: 'curvedDownArrow', 106: 'cloudCallout', 107: 'ellipseRibbon',
  108: 'ellipseRibbon2', 109: 'flowChartProcess', 110: 'flowChartDecision', 111: 'flowChartInputOutput', 112: 'flowChartPredefinedProcess',
  113: 'flowChartInternalStorage', 114: 'flowChartDocument', 115: 'flowChartMultidocument', 116: 'flowChartTerminator', 117: 'flowChartPreparation',
  118: 'flowChartManualInput', 119: 'flowChartManualOperation', 121: 'flowChartPunchedCard', 122: 'flowChartPunchedTape',
  123: 'flowChartSummingJunction', 124: 'flowChartOr', 125: 'flowChartCollate', 126: 'flowChartSort', 127: 'flowChartExtract', 128: 'flowChartMerge',
  129: 'flowChartOfflineStorage', 130: 'flowChartOnlineStorage', 131: 'flowChartMagneticTape', 132: 'flowChartMagneticDisk',
  133: 'flowChartMagneticDrum', 134: 'flowChartDisplay', 135: 'flowChartDelay', 176: 'flowChartAlternateProcess', 178: 'callout1',
  179: 'accentCallout1', 180: 'borderCallout1', 181: 'accentBorderCallout1', 182: 'leftRightUpArrow', 183: 'sun', 184: 'moon', 185: 'bracketPair',
  186: 'bracePair', 187: 'star4', 188: 'doubleWave', 189: 'actionButtonBlank', 190: 'actionButtonHome', 191: 'actionButtonHelp',
  192: 'actionButtonInformation', 193: 'actionButtonForwardNext', 194: 'actionButtonBackPrevious', 195: 'actionButtonEnd',
  196: 'actionButtonBeginning', 197: 'actionButtonReturn', 198: 'actionButtonDocument', 199: 'actionButtonSound', 200: 'actionButtonMovie',
  203: 'roundRect',
};
export const LINES = new Set([20, 32, 33, 34, 35, 36, 37, 38, 39, 40]);

/** An Office Art array (IMsoArray): its elements' bytes, each `size` long. */
export function msoArray(b) {
  if (!b || b.length < 6) return [];
  const n = u16(b, 0);
  let size = u16(b, 4);
  if (size === 0xfff0) size = 4;
  const out = [];
  for (let i = 0; i < n && 6 + (i + 1) * size <= b.length; i++) out.push(b.subarray(6 + i * size, 6 + (i + 1) * size));
  return out;
}

/**
 * A freeform's outline, as path commands in the shape's own pixels: its
 * points (pVertices) in the coordinates geoLeft..geoRight, geoTop..geoBottom
 * name, walked by its segments (pSegmentInfo) — move, line, curve, close —
 * or joined in order when it has none.
 */
export function freeformPath(props, box) {
  const vertices = msoArray(props.get(0x0145)?.complex).map((e) => (e.length >= 8 ? [i32(e, 0), i32(e, 4)] : [i16(e, 0), i16(e, 2)]));
  if (vertices.length < 2) return null;
  const left = props.get(0x0140)?.op ?? 0;
  const top = props.get(0x0141)?.op ?? 0;
  const right = props.get(0x0142)?.op ?? 21600;
  const bottom = props.get(0x0143)?.op ?? 21600;
  const sx = box.w / ((right | 0) - (left | 0) || 1);
  const sy = box.h / ((bottom | 0) - (top | 0) || 1);
  const pt = (v) => [(v[0] - (left | 0)) * sx, (v[1] - (top | 0)) * sy];
  const commands = [];
  let k = 0;
  let closed = false;
  const segments = msoArray(props.get(0x0146)?.complex).map((e) => u16(e, 0));
  if (!segments.length) {
    commands.push({ op: 'M', pts: [pt(vertices[0])] });
    for (let i = 1; i < vertices.length; i++) commands.push({ op: 'L', pts: [pt(vertices[i])] });
  } else {
    for (const seg of segments) {
      const kind = seg >> 13;
      const count = seg & 0x1fff;
      if (kind === 2 && vertices[k]) commands.push({ op: 'M', pts: [pt(vertices[k++])] });
      else if (kind === 0) for (let i = 0; i < Math.max(1, count) && vertices[k]; i++) commands.push({ op: 'L', pts: [pt(vertices[k++])] });
      else if (kind === 1) for (let i = 0; i < Math.max(1, count) && vertices[k + 2]; i++) { commands.push({ op: 'C', pts: [pt(vertices[k]), pt(vertices[k + 1]), pt(vertices[k + 2])] }); k += 3; }
      else if (kind === 3) { commands.push({ op: 'Z' }); closed = true; }
      else if (kind === 4) break;
      else if (kind === 5) k += seg & 0xff; // an escape: its points skipped
    }
  }
  if (!commands.length || commands[0].op !== 'M') return null;
  return { commands, w: box.w, h: box.h, filled: closed };
}

/* ── fills, outlines, shadows ─────────────────────────────────────────── */

/** An Office Art 16.16 fraction (an opacity) as the thousandths PowerPoint means: 0x9999 is 60%. */
const fraction = (op) => Math.max(0, Math.min(1, Math.round(((op >>> 0) / 65536) * 1000) / 1000));

/**
 * A shape's fill as a colour, or 'none'. What a shape does not say is
 * Office Art's own default — filled, white — not the drawing group's
 * defaults, which are only what PowerPoint gives a new shape.
 */
export function artFill(props, scheme, background = false) {
  const get = (id) => props.get(id);
  const bools = get(0x01bf)?.op;
  const filled = bools != null && bools & 0x100000 ? Boolean(bools & 0x10) : true;
  if (!filled) return 'none';
  const type = get(0x0180)?.op ?? 0;
  const colour = get(0x0181) ? artColour(get(0x0181).op, scheme) : background ? scheme[0] : 'FFFFFF';
  // A shaded fill is a gradient; a texture is drawn as its colour; a picture fill as nothing this can draw.
  if (type === 3) return background ? null : 'none';
  // How opaque each colour is, where it is not wholly.
  const opacity = (id) => (get(id) ? fraction(get(id).op) : 1);
  const tint = (hex, a) => (a < 1 ? { color: '#' + hex, alpha: a } : '#' + hex);
  if (type >= 4 && type <= 8 && colour) {
    // The second colour, white where it does not say (Office Art's own).
    const back = get(0x0183) ? artColour(get(0x0183).op, scheme) : 'FFFFFF';
    if (back) return gradientOf(props, tint(colour, opacity(0x0182)), tint(back, opacity(0x0184)), type);
  }
  return colour ? tint(colour, opacity(0x0182)) : 'none';
}

/**
 * Office Art's shaded fill as a linear gradient: { gradient: { stops,
 * angle } }, the angle DrawingML's. Which colour comes first follows
 * LibreOffice's reading — the angle's sign, the focus (none, negative or
 * about half, which makes it run out and back) and a centre or shape-shaded
 * fill each turn it round; a centre or shape-shaded fill is drawn linear.
 */
function gradientOf(props, fore, back, type) {
  const raw = (props.get(0x018b)?.op ?? 0) | 0;
  const focus = (props.get(0x018c)?.op ?? 0) | 0;
  const axial = Math.abs(focus) > 40 && Math.abs(focus) < 60;
  let swap = raw >= 0;
  if (!focus || focus < 0) swap = !swap;
  if (axial) swap = !swap;
  if (type === 5 || type === 6) swap = !swap;
  const [start, end] = swap ? [fore, back] : [back, fore];
  const stop = (pos, c) => (typeof c === 'string' ? { pos, color: c } : { pos, ...c });
  const stops = axial ? [stop(0, start), stop(0.5, end), stop(1, start)] : [stop(0, start), stop(1, end)];
  return { gradient: { stops, angle: (((450 - raw / 65536) % 360) + 360) % 360 } };
}

/** A shape's outline, or 'none': black, three quarters of a point, where it does not say. */
export function artLine(props, scheme, defaultOn) {
  const get = (id) => props.get(id);
  const bools = get(0x01ff)?.op;
  const on = bools != null && bools & 0x80000 ? Boolean(bools & 0x08) : defaultOn;
  if (!on) return 'none';
  const colour = get(0x01c0) ? artColour(get(0x01c0).op, scheme) : '000000';
  return { color: '#' + (colour || '000000'), width: Math.max(0.25, (get(0x01cb)?.op ?? 9525) / 12700) };
}

/** A shape's shadow: Office Art's offset, colour and opacity; it keeps no blur, so PowerPoint's own five points. Null for none. */
export function artShadow(props, scheme) {
  const bits = props.get(0x023f)?.op ?? 0;
  if (!(bits & 0x20000 && bits & 0x2)) return null;
  const dx = (props.get(0x0205)?.op ?? 25400) | 0;
  const dy = (props.get(0x0206)?.op ?? 25400) | 0;
  const tone = props.get(0x0201) ? artColour(props.get(0x0201).op, scheme) : '808080';
  return {
    color: '#' + (tone || '808080'), alpha: props.has(0x0204) ? fraction(props.get(0x0204).op) : 1,
    dist: Math.hypot(dx, dy) / 12700, dir: ((Math.atan2(dy, dx) * 180) / Math.PI + 360) % 360, blur: 5,
  };
}
