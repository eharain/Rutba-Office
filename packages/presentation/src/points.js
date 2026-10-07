// Edit Points: a shape's outline as points a person can drag.
//
// PowerPoint turns a shape into a freeform the moment its points are edited:
// the preset's outline becomes a path of the shape's own — <a:custGeom> —
// and from then on that path is what is drawn and saved. Here a path is a
// list of commands in the shape's own box, in pixels: M (start a figure),
// L (a straight line to a point), C (a curve: two control points and the
// point it ends on) and Z (close the figure). A quadratic curve or an arc
// read from a file becomes cubic curves, which is how PowerPoint's own
// editor shows them.
//
// Pure: commands in, commands out. Nothing here knows a slide or a file.

/** A quarter circle's control points sit this far along its tangents, as a share of the radius. */
const K = 0.5522847498;

const L = (x, y) => ({ op: 'L', pts: [[x, y]] });
const M = (x, y) => ({ op: 'M', pts: [[x, y]] });
const C = (a, b, c) => ({ op: 'C', pts: [a, b, c] });
const Z = () => ({ op: 'Z', pts: [] });
const polygon = (pts) => [M(...pts[0]), ...pts.slice(1).map((p) => L(...p)), Z()];

/** An ellipse in a box, as four quarter curves from its top. */
function ellipse(x, y, w, h) {
  const cx = x + w / 2;
  const cy = y + h / 2;
  const rx = w / 2;
  const ry = h / 2;
  return [
    M(cx, y),
    C([cx + rx * K, y], [x + w, cy - ry * K], [x + w, cy]),
    C([x + w, cy + ry * K], [cx + rx * K, y + h], [cx, y + h]),
    C([cx - rx * K, y + h], [x, cy + ry * K], [x, cy]),
    C([x, cy - ry * K], [cx - rx * K, y], [cx, y]),
    Z(),
  ];
}

/**
 * A preset's outline as commands in a w × h box — the same outline the
 * slide draws for it, so turning it into points changes nothing on the
 * page. A preset the page draws as a rectangle comes out as one.
 */
export function presetCommands(preset, w, h) {
  const r = Math.min(w, h);
  switch (preset) {
    case 'ellipse':
    case 'circle':
      return ellipse(0, 0, w, h);
    case 'roundRect': {
      const c = r * 0.16;
      const k = c * (1 - K);
      return [
        M(c, 0), L(w - c, 0), C([w - k, 0], [w, k], [w, c]),
        L(w, h - c), C([w, h - k], [w - k, h], [w - c, h]),
        L(c, h), C([k, h], [0, h - k], [0, h - c]),
        L(0, c), C([0, k], [k, 0], [c, 0]), Z(),
      ];
    }
    case 'triangle':
      return polygon([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle':
      return polygon([[0, 0], [0, h], [w, h]]);
    case 'diamond':
      return polygon([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram':
      return polygon([[w * 0.25, 0], [w, 0], [w * 0.75, h], [0, h]]);
    case 'trapezoid':
      return polygon([[w * 0.25, 0], [w * 0.75, 0], [w, h], [0, h]]);
    case 'pentagon':
    case 'hexagon':
    case 'octagon': {
      const sides = { pentagon: 5, hexagon: 6, octagon: 8 }[preset];
      return polygon(Array.from({ length: sides }, (_, i) => {
        const a = (Math.PI * 2 * i) / sides - Math.PI / 2;
        return [w / 2 + (Math.cos(a) * w) / 2, h / 2 + (Math.sin(a) * h) / 2];
      }));
    }
    case 'star5':
      return polygon(Array.from({ length: 10 }, (_, i) => {
        const rad = i % 2 ? 0.4 : 1;
        const a = (Math.PI * i) / 5 - Math.PI / 2;
        return [w / 2 + (Math.cos(a) * w * rad) / 2, h / 2 + (Math.sin(a) * h * rad) / 2];
      }));
    case 'rightArrow':
      return polygon([[0, h * 0.3], [w * 0.6, h * 0.3], [w * 0.6, 0], [w, h / 2], [w * 0.6, h], [w * 0.6, h * 0.7], [0, h * 0.7]]);
    case 'chevron':
      return polygon([[0, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [0, h], [w * 0.25, h / 2]]);
    case 'line':
    case 'straightConnector1':
      return [M(0, 0), L(w, h)];
    default:
      return polygon([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

/**
 * An SVG arc (endpoint form) as cubic curves, a quarter turn or less each —
 * the conversion every SVG renderer does, after the specification's own
 * appendix (F.6).
 */
function arcToCurves([x1, y1], rx, ry, angle, large, sweep, [x2, y2]) {
  if (!rx || !ry) return [L(x2, y2)];
  const phi = (angle * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const xp = cos * dx + sin * dy;
  const yp = -sin * dx + cos * dy;
  rx = Math.abs(rx);
  ry = Math.abs(ry);
  const scale = (xp * xp) / (rx * rx) + (yp * yp) / (ry * ry);
  if (scale > 1) {
    rx *= Math.sqrt(scale);
    ry *= Math.sqrt(scale);
  }
  const num = rx * rx * ry * ry - rx * rx * yp * yp - ry * ry * xp * xp;
  const den = rx * rx * yp * yp + ry * ry * xp * xp;
  const co = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, num / den));
  const cxp = (co * rx * yp) / ry;
  const cyp = (-co * ry * xp) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angleOf = (ux, uy, vx, vy) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const t1 = angleOf(1, 0, (xp - cxp) / rx, (yp - cyp) / ry);
  let dt = angleOf((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry);
  if (!sweep && dt > 0) dt -= Math.PI * 2;
  if (sweep && dt < 0) dt += Math.PI * 2;
  const parts = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / parts;
  const k = (4 / 3) * Math.tan(step / 4);
  const point = (t) => [cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
  const tangent = (t) => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
  const out = [];
  for (let i = 0; i < parts; i++) {
    const a = t1 + i * step;
    const b = a + step;
    const [p0x, p0y] = point(a);
    const [p3x, p3y] = i === parts - 1 ? [x2, y2] : point(b);
    const [d0x, d0y] = tangent(a);
    const [d3x, d3y] = tangent(b);
    out.push(C([p0x + k * d0x, p0y + k * d0y], [p3x - k * d3x, p3y - k * d3y], [p3x, p3y]));
  }
  return out;
}

/**
 * SVG path data — absolute M, L, C, Q, A and Z, which is what the deck's
 * reader writes for a custom geometry — as commands, scaled from the path's
 * own `w × h` box onto a `width × height` one.
 */
export function parsePath(d, { w = 1, h = 1, width = w, height = h } = {}) {
  const sx = width / (w || 1);
  const sy = height / (h || 1);
  const out = [];
  let at = [0, 0];
  let start = [0, 0];
  for (const [, op, args] of String(d || '').matchAll(/([MLCQAZ])([^MLCQAZ]*)/gi)) {
    const nums = args.trim().split(/[\s,]+/).filter((v) => v !== '').map(Number);
    const p = (i) => [nums[i] * sx, nums[i + 1] * sy];
    const o = op.toUpperCase();
    if (o === 'M') {
      for (let i = 0; i + 1 < nums.length; i += 2) {
        out.push(i === 0 ? M(...p(i)) : L(...p(i)));
        at = p(i);
        if (i === 0) start = at;
      }
    } else if (o === 'L') {
      for (let i = 0; i + 1 < nums.length; i += 2) { out.push(L(...p(i))); at = p(i); }
    } else if (o === 'C') {
      for (let i = 0; i + 5 < nums.length; i += 6) { out.push(C(p(i), p(i + 2), p(i + 4))); at = p(i + 4); }
    } else if (o === 'Q') {
      for (let i = 0; i + 3 < nums.length; i += 4) {
        const q = p(i);
        const e = p(i + 2);
        out.push(C([at[0] + (2 / 3) * (q[0] - at[0]), at[1] + (2 / 3) * (q[1] - at[1])], [e[0] + (2 / 3) * (q[0] - e[0]), e[1] + (2 / 3) * (q[1] - e[1])], e));
        at = e;
      }
    } else if (o === 'A') {
      for (let i = 0; i + 6 < nums.length; i += 7) {
        const e = p(i + 5);
        out.push(...arcToCurves(at, nums[i] * sx, nums[i + 1] * sy, nums[i + 2], nums[i + 3], nums[i + 4], e));
        at = e;
      }
    } else if (o === 'Z') {
      out.push(Z());
      at = start;
    }
  }
  return out;
}

const r2 = (v) => Math.round(v * 100) / 100;

/** Commands as SVG path data. */
export function pathData(commands) {
  return commands.map((c) => (c.op === 'Z' ? 'Z' : c.op + c.pts.map(([x, y]) => `${r2(x)} ${r2(y)}`).join(' '))).join(' ');
}

const same = (a, b) => Math.abs(a[0] - b[0]) < 0.01 && Math.abs(a[1] - b[1]) < 0.01;

/** Where each figure starts and ends: [startIndex, endIndex (exclusive), closed]. */
function figures(commands) {
  const out = [];
  let begin = -1;
  commands.forEach((c, i) => {
    if (c.op === 'M') {
      if (begin >= 0) out.push([begin, i, false]);
      begin = i;
    } else if (c.op === 'Z' && begin >= 0) {
      out.push([begin, i + 1, true]);
      begin = -1;
    }
  });
  if (begin >= 0) out.push([begin, commands.length, false]);
  return out;
}

/**
 * The points a person can drag: the end of each M, L and C. A closed
 * figure whose last command lands back on its start (an ellipse's last
 * quarter) shows that place once.
 */
export function anchors(commands) {
  const out = [];
  for (const [begin, end, closed] of figures(commands)) {
    const last = commands[end - 1]?.op === 'Z' ? end - 2 : end - 1;
    for (let i = begin; i <= last; i++) {
      const c = commands[i];
      if (c.op === 'Z') continue;
      const p = c.pts[c.pts.length - 1];
      if (closed && i === last && i !== begin && same(p, commands[begin].pts[0])) continue;
      out.push({ index: i, x: p[0], y: p[1] });
    }
  }
  return out;
}

const copy = (commands) => commands.map((c) => ({ op: c.op, pts: c.pts.map((p) => [...p]) }));

/** The command that lands back on figure `begin`'s start, when a closed figure has one. */
function closingTwin(commands, index) {
  const fig = figures(commands).find(([b]) => b === index);
  if (!fig || !fig[2]) return -1;
  const last = fig[1] - 2;
  return last > index && same(commands[last].pts[commands[last].pts.length - 1], commands[index].pts[0]) ? last : -1;
}

/**
 * The point at command `index` moved to (x, y), the curves either side of
 * it bending with it: the control point that leads into it and the one that
 * leads out move by the same amount, as PowerPoint's own handles do.
 */
export function moveAnchor(commands, index, x, y) {
  const out = copy(commands);
  const c = out[index];
  if (!c || c.op === 'Z') return out;
  const end = c.pts.length - 1;
  const dx = x - c.pts[end][0];
  const dy = y - c.pts[end][1];
  const shift = (p) => { p[0] += dx; p[1] += dy; };
  shift(c.pts[end]);
  if (c.op === 'C') shift(c.pts[1]);
  const next = out[index + 1];
  if (next?.op === 'C') shift(next.pts[0]);
  // A figure's start, and the command that comes back to it.
  const twin = c.op === 'M' ? closingTwin(commands, index) : -1;
  if (twin >= 0) {
    const t = out[twin];
    shift(t.pts[t.pts.length - 1]);
    if (t.op === 'C') shift(t.pts[1]);
  }
  return out;
}

/**
 * The point at command `index` taken out, the figure joining its neighbours
 * directly. Null when that would leave a figure of fewer than two points
 * (or a closed one of fewer than three) — there is nothing left to draw.
 */
export function deleteAnchor(commands, index) {
  const fig = figures(commands).find(([b, e]) => index >= b && index < e);
  if (!fig) return null;
  const count = anchors(commands.slice(fig[0], fig[1])).length;
  if (count <= (fig[2] ? 3 : 2)) return null;
  const out = copy(commands);
  const twin = out[index].op === 'M' ? closingTwin(commands, index) : -1;
  if (out[index].op === 'M') {
    // The next point starts the figure; a closed figure's way back comes to it.
    const next = out[index + 1];
    out[index + 1] = M(...next.pts[next.pts.length - 1]);
    if (twin >= 0) {
      const t = out[twin];
      t.pts[t.pts.length - 1] = [...out[index + 1].pts[0]];
    }
  }
  out.splice(index, 1);
  return out;
}

/** The point a segment passes halfway along, and the two segments it splits into. */
function split(from, c) {
  if (c.op === 'C') {
    const [p1, p2, p3] = c.pts;
    const mid = (a, b) => [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
    const a = mid(from, p1);
    const b = mid(p1, p2);
    const d = mid(p2, p3);
    const e = mid(a, b);
    const f = mid(b, d);
    const g = mid(e, f);
    return [C(a, e, g), C(f, d, p3)];
  }
  const p = c.pts[0];
  return [L((from[0] + p[0]) / 2, (from[1] + p[1]) / 2), L(...p)];
}

/**
 * A point added halfway along the segment that ends at command `index` —
 * or, for a Z, along the line that closes the figure.
 */
export function insertAnchor(commands, index) {
  const out = copy(commands);
  const c = out[index];
  if (!c || index === 0 || c.op === 'M') return out;
  const prev = out[index - 1];
  const from = prev.op === 'Z' ? null : prev.pts[prev.pts.length - 1];
  if (!from) return out;
  if (c.op === 'Z') {
    const fig = figures(commands).find(([, e]) => e === index + 1);
    const start = fig ? out[fig[0]].pts[0] : from;
    out.splice(index, 0, L((from[0] + start[0]) / 2, (from[1] + start[1]) / 2));
    return out;
  }
  out.splice(index, 1, ...split(from, c));
  return out;
}

/**
 * The segment nearest a point, by the command it ends at — what a
 * right-click on the outline adds a point to.
 */
export function nearestSegment(commands, x, y) {
  let best = { index: -1, distance: Infinity };
  let at = null;
  let start = null;
  commands.forEach((c, i) => {
    const sample = (pts) => {
      for (const [px, py] of pts) {
        const dd = Math.hypot(px - x, py - y);
        if (dd < best.distance) best = { index: i, distance: dd };
      }
    };
    if (c.op === 'M') {
      at = c.pts[0];
      start = at;
      return;
    }
    if (!at) return;
    const to = c.op === 'Z' ? start : c.pts[c.pts.length - 1];
    const pts = [];
    for (let s = 1; s < 16; s++) {
      const t = s / 16;
      if (c.op === 'C') {
        const [p1, p2, p3] = c.pts;
        const u = 1 - t;
        pts.push([u * u * u * at[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0], u * u * u * at[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1]]);
      } else pts.push([at[0] + (to[0] - at[0]) * t, at[1] + (to[1] - at[1]) * t]);
    }
    sample(pts);
    at = to;
  });
  return best.index;
}

/**
 * The box the path fills — its points, and its curves sampled, since a
 * curve's control points lie outside what it draws — and the path moved
 * so the box starts at 0, 0: after points are dragged outside the shape,
 * the shape is the path's box, as PowerPoint makes it.
 */
export function fit(commands) {
  const xs = [];
  const ys = [];
  let at = null;
  for (const c of commands) {
    if (c.op === 'Z') continue;
    if (c.op === 'C' && at) {
      const [p1, p2, p3] = c.pts;
      for (let s = 0; s <= 16; s++) {
        const t = s / 16;
        const u = 1 - t;
        xs.push(u * u * u * at[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0]);
        ys.push(u * u * u * at[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1]);
      }
    } else for (const [x, y] of c.pts) { xs.push(x); ys.push(y); }
    at = c.pts[c.pts.length - 1];
  }
  if (!xs.length) return { commands, dx: 0, dy: 0, w: 1, h: 1 };
  const minX = Math.min(...xs);
  const minY = Math.min(...ys);
  const w = Math.max(1, Math.max(...xs) - minX);
  const h = Math.max(1, Math.max(...ys) - minY);
  return { commands: commands.map((c) => ({ op: c.op, pts: c.pts.map(([x, y]) => [x - minX, y - minY]) })), dx: minX, dy: minY, w, h };
}

/** EMU in a pixel. */
const EMU = 9525;

/**
 * Commands as DrawingML's own custom geometry, in a path box the shape's
 * size in EMU so nothing is rounded away. `filled` false is a line: its
 * figures drawn, never filled.
 */
export function custGeomXml(commands, w, h, { filled = true } = {}) {
  const e = (v) => Math.round(v * EMU);
  const pt = ([x, y]) => `<a:pt x="${e(x)}" y="${e(y)}"/>`;
  const body = commands.map((c) => {
    if (c.op === 'M') return `<a:moveTo>${pt(c.pts[0])}</a:moveTo>`;
    if (c.op === 'L') return `<a:lnTo>${pt(c.pts[0])}</a:lnTo>`;
    if (c.op === 'C') return `<a:cubicBezTo>${c.pts.map(pt).join('')}</a:cubicBezTo>`;
    return '<a:close/>';
  }).join('');
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="${Math.max(1, e(w))}" h="${Math.max(1, e(h))}"${filled ? '' : ' fill="none"'}>${body}</a:path></a:pathLst></a:custGeom>`;
}
