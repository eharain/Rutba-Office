// OpenDocument custom shapes: <draw:enhanced-geometry> worked out into the
// outline it draws.
//
// A shape LibreOffice, Excel or PowerPoint writes to ODF is mostly a
// "non-primitive" custom shape: its outline is draw:enhanced-path, a path in
// the coordinates of svg:viewBox whose numbers may be the shape's
// modifiers ($0, $1 …) or its equations (?f0, ?f1 …), each equation a
// draw:formula over the modifiers, other equations, the view box's
// left/top/right/bottom/width/height, the shape's logical size
// (logwidth/logheight), pi and the functions abs, sqrt, sin, cos, tan,
// atan, atan2, min, max and if. The path's commands ([ODF 1.2] 19.145):
//
//   M moveto, L lineto, C curveto, Q quadratic curveto, Z close, N end of
//   a figure, F not filled, S not stroked; T/U an ellipse's arc by centre,
//   radii and angles (U starting a figure); A/B (counter-clockwise) and
//   W/V (clockwise) an ellipse's arc by its bounding box and two points
//   (B and V starting a figure); X/Y a quarter ellipse leaving across or
//   down; G an arc by radii, start and swing, as DrawingML's arcTo.
//
// Out come the figures as moveto/lineto/curveto commands in the shape's
// own box (pixels), each with whether it is filled and stroked — an arc
// as cubic Béziers — for DrawingML's custom geometry. A path that cannot
// be worked out gives null, and the shape keeps the preset it maps to.

const KAPPA = 0.5522847498;

/** A formula's tokens: numbers, names, ?equations, $modifiers, operators. */
function tokens(text) {
  const out = [];
  const re = /\s*(?:(\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+)|(\?[A-Za-z0-9_]+)|(\$\d+)|([A-Za-z][A-Za-z0-9_]*)|([-+*/(),]))/y;
  let m;
  let at = 0;
  while (at < text.length) {
    re.lastIndex = at;
    m = re.exec(text);
    if (!m) {
      if (/^\s*$/.test(text.slice(at))) break;
      throw new Error(`cannot read "${text.slice(at, at + 12)}"`);
    }
    out.push(m[1] != null ? { n: Number(m[1]) } : m[2] ? { eq: m[2].slice(1) } : m[3] ? { mod: Number(m[3].slice(1)) } : m[4] ? { id: m[4].toLowerCase() } : { op: m[5] });
    at = re.lastIndex;
  }
  return out;
}

const FUNCTIONS = {
  abs: (a) => Math.abs(a), sqrt: (a) => Math.sqrt(Math.max(0, a)), sin: (a) => Math.sin(a), cos: (a) => Math.cos(a), tan: (a) => Math.tan(a),
  atan: (a) => Math.atan(a), atan2: (y, x) => Math.atan2(y, x), min: (a, b) => Math.min(a, b), max: (a, b) => Math.max(a, b),
  if: (c, a, b) => (c > 0 ? a : b),
};

/**
 * The values a shape's names have, and its equations worked out on demand —
 * each once, a loop between them taken as nought.
 */
function evaluator(node, box) {
  const vb = String(node.attrs['svg:viewBox'] || '0 0 21600 21600').trim().split(/[\s,]+/).map(Number);
  const [left, top, width, height] = vb.length === 4 && vb.every(Number.isFinite) && vb[2] > 0 && vb[3] > 0 ? vb : [0, 0, 21600, 21600];
  const modifiers = String(node.attrs['draw:modifiers'] || '').trim().split(/[\s,]+/).filter(Boolean).map(Number);
  const formulas = new Map();
  for (const c of node.children || []) {
    if (typeof c !== 'string' && c.name === 'draw:equation') formulas.set(c.attrs['draw:name'], c.attrs['draw:formula'] || '0');
  }
  // The shape's logical size: only its proportion matters, as width against height.
  const names = {
    pi: Math.PI, left, top, right: left + width, bottom: top + height, width, height,
    logwidth: Math.max(1, box.width * 26.458), logheight: Math.max(1, box.height * 26.458),
    xstretch: 0, ystretch: 0, hasstroke: 1, hasfill: 1,
  };
  const done = new Map();
  const busy = new Set();
  const equation = (name) => {
    if (done.has(name)) return done.get(name);
    if (busy.has(name) || !formulas.has(name)) return 0;
    busy.add(name);
    const v = run(tokens(formulas.get(name)));
    busy.delete(name);
    done.set(name, Number.isFinite(v) ? v : 0);
    return done.get(name);
  };
  const run = (list) => {
    let i = 0;
    const peek = () => list[i];
    const expr = () => {
      let v = term();
      while (peek()?.op === '+' || peek()?.op === '-') v = list[i++].op === '+' ? v + term() : v - term();
      return v;
    };
    const term = () => {
      let v = factor();
      while (peek()?.op === '*' || peek()?.op === '/') {
        const op = list[i++].op;
        const r = factor();
        v = op === '*' ? v * r : r === 0 ? 0 : v / r;
      }
      return v;
    };
    const factor = () => {
      const t = list[i++];
      if (!t) throw new Error('a formula ends early');
      if (t.op === '-') return -factor();
      if (t.op === '+') return factor();
      if (t.op === '(') {
        const v = expr();
        if (list[i]?.op === ')') i++;
        return v;
      }
      if (t.n != null) return t.n;
      if (t.eq != null) return equation(t.eq);
      if (t.mod != null) return modifiers[t.mod] ?? 0;
      if (t.id) {
        if (FUNCTIONS[t.id] && list[i]?.op === '(') {
          i++;
          const args = [];
          if (list[i]?.op !== ')') {
            args.push(expr());
            while (list[i]?.op === ',') { i++; args.push(expr()); }
          }
          if (list[i]?.op === ')') i++;
          return FUNCTIONS[t.id](...args);
        }
        if (t.id in names) return names[t.id];
        throw new Error(`no value for "${t.id}"`);
      }
      throw new Error(`cannot read "${t.op}"`);
    };
    const v = expr();
    return v;
  };
  const value = (token) => {
    if (/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(token)) return Number(token);
    if (token.startsWith('?')) return equation(token.slice(1));
    if (token.startsWith('$')) return modifiers[Number(token.slice(1))] ?? 0;
    const lower = token.toLowerCase();
    if (lower in names) return names[lower];
    throw new Error(`no value for "${token}"`);
  };
  return { value, left, top, width, height };
}

/** An ODF angle on a path: degrees, or — written as Office's 16.16 fixed point — that over 65536. */
const degrees = (v) => (Math.abs(v) >= 65536 ? v / 65536 : v);

/**
 * Cubic Béziers along an ellipse (centre cx, cy, radii rx, ry) from screen
 * angle a0 to a1 (radians, clockwise on screen positive), as points after
 * the start: [c1, c2, end] per quarter at most.
 */
function arcCurves(cx, cy, rx, ry, a0, a1) {
  const out = [];
  const n = Math.max(1, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 2) - 1e-9));
  const step = (a1 - a0) / n;
  const k = (4 / 3) * Math.tan(step / 4);
  for (let s = 0; s < n; s++) {
    const t0 = a0 + s * step;
    const t1 = t0 + step;
    const p0 = [cx + rx * Math.cos(t0), cy + ry * Math.sin(t0)];
    const p1 = [cx + rx * Math.cos(t1), cy + ry * Math.sin(t1)];
    out.push([
      [p0[0] - k * rx * Math.sin(t0), p0[1] + k * ry * Math.cos(t0)],
      [p1[0] + k * rx * Math.sin(t1), p1[1] - k * ry * Math.cos(t1)],
      p1,
    ]);
  }
  return out;
}

/** The parametric angle of a point seen from an ellipse's centre, its radii taken into account. */
const angleOn = (cx, cy, rx, ry, x, y) => Math.atan2((y - cy) / (ry || 1), (x - cx) / (rx || 1));

/**
 * The figures an enhanced geometry draws, in a box `width` × `height`
 * pixels: [{ commands: [{ op: 'M'|'L'|'C'|'Z', pts }], fill, stroke }] —
 * or null when its path cannot be worked out.
 */
export function enhancedFigures(node, { width, height }) {
  const pathText = node?.attrs?.['draw:enhanced-path'];
  if (!pathText || !(width > 0) || !(height > 0)) return null;
  try {
    const ev = evaluator(node, { width, height });
    const flipX = node.attrs['draw:mirror-horizontal'] === 'true';
    const flipY = node.attrs['draw:mirror-vertical'] === 'true';
    const sx = (x) => { const v = ((x - ev.left) / ev.width) * width; return flipX ? width - v : v; };
    const sy = (y) => { const v = ((y - ev.top) / ev.height) * height; return flipY ? height - v : v; };
    const list = pathText.trim().split(/[\s,]+/).filter(Boolean);
    const figures = [];
    let figure = null;
    let fill = true;
    let stroke = true;
    let at = null;
    let start = null;
    const begin = () => {
      if (!figure) figure = { commands: [], fill, stroke };
      return figure;
    };
    const finish = () => {
      if (figure && figure.commands.some((c) => c.op !== 'M')) figures.push(figure);
      figure = null;
    };
    const move = (p) => { if (figure?.commands.length) { /* a moveto inside a figure starts a sub-path of it */ } begin().commands.push({ op: 'M', pts: [p] }); at = p; start = p; };
    const line = (p) => { if (!at) return move(p); begin().commands.push({ op: 'L', pts: [p] }); at = p; };
    const curve = (c1, c2, p) => { if (!at) move(c1); begin().commands.push({ op: 'C', pts: [c1, c2, p] }); at = p; };
    // An arc in view-box coordinates, its points scaled into the box.
    const arc = (cx, cy, rx, ry, a0, a1, joined) => {
      const first = [cx + rx * Math.cos(a0), cy + ry * Math.sin(a0)];
      if (joined && at) line([sx(first[0]), sy(first[1])]);
      else move([sx(first[0]), sy(first[1])]);
      for (const [c1, c2, p] of arcCurves(cx, cy, rx, ry, a0, a1)) curve([sx(c1[0]), sy(c1[1])], [sx(c2[0]), sy(c2[1])], [sx(p[0]), sy(p[1])]);
    };
    let i = 0;
    let cmd = null;
    let quadrant = null;
    const isCommand = (t) => /^[A-Za-z]$/.test(t) && !/^(pi|left|top|right|bottom|width|height)$/i.test(t);
    const take = (n) => {
      const out = [];
      for (let k = 0; k < n; k++) {
        if (i >= list.length || isCommand(list[i])) throw new Error('a path command is short of numbers');
        out.push(ev.value(list[i++]));
      }
      return out;
    };
    let guard = 0;
    while (i < list.length && guard++ < 100000) {
      if (isCommand(list[i])) {
        cmd = list[i++].toUpperCase();
        if (cmd === 'X' || cmd === 'Y') quadrant = cmd;
        if (cmd === 'Z') { if (figure) figure.commands.push({ op: 'Z', pts: [] }); at = start; continue; }
        if (cmd === 'N') { finish(); fill = true; stroke = true; at = null; continue; }
        if (cmd === 'F') { fill = false; if (figure) figure.fill = false; continue; }
        if (cmd === 'S') { stroke = false; if (figure) figure.stroke = false; continue; }
        if (i >= list.length || isCommand(list[i])) continue;
      }
      if (!cmd) { i++; continue; }
      switch (cmd) {
        case 'M': { const [x, y] = take(2); move([sx(x), sy(y)]); cmd = 'L'; break; }
        case 'L': { const [x, y] = take(2); line([sx(x), sy(y)]); break; }
        case 'C': { const [x1, y1, x2, y2, x, y] = take(6); curve([sx(x1), sy(y1)], [sx(x2), sy(y2)], [sx(x), sy(y)]); break; }
        case 'Q': {
          const [qx, qy, x, y] = take(4);
          const p0 = at || [sx(qx), sy(qy)];
          const q = [sx(qx), sy(qy)];
          const p = [sx(x), sy(y)];
          curve([p0[0] + (2 / 3) * (q[0] - p0[0]), p0[1] + (2 / 3) * (q[1] - p0[1])], [p[0] + (2 / 3) * (q[0] - p[0]), p[1] + (2 / 3) * (q[1] - p[1])], p);
          break;
        }
        case 'T':
        case 'U': {
          // Centre, radii, and angles counter-clockwise from three o'clock.
          const [cx, cy, rx, ry, t0, t1] = take(6);
          let a0 = (-degrees(t0) * Math.PI) / 180;
          let a1 = (-degrees(t1) * Math.PI) / 180;
          if (a1 >= a0) a1 -= Math.ceil((a1 - a0) / (2 * Math.PI) + 1e-9) * 2 * Math.PI;
          if (Math.abs(a1 - a0) < 1e-9) a1 = a0 - 2 * Math.PI;
          arc(cx, cy, Math.abs(rx), Math.abs(ry), a0, a1, cmd === 'T');
          break;
        }
        case 'A':
        case 'B':
        case 'W':
        case 'V': {
          const [x1, y1, x2, y2, x3, y3, x4, y4] = take(8);
          const cx = (x1 + x2) / 2;
          const cy = (y1 + y2) / 2;
          const rx = Math.abs(x2 - x1) / 2;
          const ry = Math.abs(y2 - y1) / 2;
          if (rx < 1e-9 || ry < 1e-9) break; // a box with no size draws nothing
          let a0 = angleOn(cx, cy, rx, ry, x3, y3);
          let a1 = angleOn(cx, cy, rx, ry, x4, y4);
          const clockwise = cmd === 'W' || cmd === 'V';
          if (clockwise) { if (a1 <= a0 + 1e-9) a1 += 2 * Math.PI; } else if (a1 >= a0 - 1e-9) a1 -= 2 * Math.PI;
          arc(cx, cy, rx, ry, a0, a1, cmd === 'A' || cmd === 'W');
          break;
        }
        case 'X':
        case 'Y': {
          // A quarter ellipse to the point, leaving across (X) or down (Y); the next turns the other way.
          const [x, y] = take(2);
          const p = [sx(x), sy(y)];
          const p0 = at || p;
          if (quadrant === 'X') curve([p0[0] + KAPPA * (p[0] - p0[0]), p0[1]], [p[0], p[1] - KAPPA * (p[1] - p0[1])], p);
          else curve([p0[0], p0[1] + KAPPA * (p[1] - p0[1])], [p[0] - KAPPA * (p[0] - p0[0]), p[1]], p);
          quadrant = quadrant === 'X' ? 'Y' : 'X';
          break;
        }
        case 'G': {
          // DrawingML's arcTo: the current point on the ellipse at the start angle, swung clockwise by the swing.
          const [wr, hr, st, sw] = take(4);
          if (!at) break;
          const a0 = (st * Math.PI) / 180;
          const a1 = a0 + (sw * Math.PI) / 180;
          // Back from the box to view-box units, where the radii are.
          const ux = flipX ? width - at[0] : at[0];
          const uy = flipY ? height - at[1] : at[1];
          const vx = ev.left + (ux / width) * ev.width;
          const vy = ev.top + (uy / height) * ev.height;
          const cx = vx - wr * Math.cos(a0);
          const cy = vy - hr * Math.sin(a0);
          for (const [c1, c2, p] of arcCurves(cx, cy, wr, hr, a0, a1)) curve([sx(c1[0]), sy(c1[1])], [sx(c2[0]), sy(c2[1])], [sx(p[0]), sy(p[1])]);
          break;
        }
        default:
          // A command this does not draw: its numbers are passed over.
          i++;
      }
    }
    finish();
    return figures.length ? figures : null;
  } catch {
    return null;
  }
}

/**
 * The figures as DrawingML's custom geometry, each figure a path in a box
 * the shape's size in EMU: `<a:custGeom>` with its pathLst.
 */
export function figuresGeometryXml(figures, width, height) {
  const e = (v) => Math.round(v * 9525);
  const pt = ([x, y]) => `<a:pt x="${e(x)}" y="${e(y)}"/>`;
  const paths = figures.map((f) => {
    const body = f.commands.map((c) => (c.op === 'M' ? `<a:moveTo>${pt(c.pts[0])}</a:moveTo>`
      : c.op === 'L' ? `<a:lnTo>${pt(c.pts[0])}</a:lnTo>`
        : c.op === 'C' ? `<a:cubicBezTo>${c.pts.map(pt).join('')}</a:cubicBezTo>`
          : '<a:close/>')).join('');
    return `<a:path w="${Math.max(1, e(width))}" h="${Math.max(1, e(height))}"${f.fill ? '' : ' fill="none"'}${f.stroke ? '' : ' stroke="0"'}>${body}</a:path>`;
  }).join('');
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
}
