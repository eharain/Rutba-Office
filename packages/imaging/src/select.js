// Selections and painted masks (the Image app's advanced editor,
// docs/STUDIO.md 7.2, I10 and I11).
//
// A selection is an alpha bitmap the size of the canvas: `{ width, height,
// data }`, `data` a Uint8Array of 0 (outside) to 255 (inside), so a feathered
// edge is the numbers in between and nothing else. A layer's painted mask is
// the same kind of bitmap. Everything here makes, combines or reshapes one:
// rectangles, ellipses and lassos rasterised with smooth edges, the magic
// wand and select by colour, feather, grow, shrink, invert, and the outline
// the window draws its marching ants along.
//
// Pure arithmetic over arrays: no DOM, no canvas. It runs in the image worker
// for anything the size of a picture, and in Node for the tests.

const clamp = (v, lo, hi) => (v < lo ? lo : v > hi ? hi : v);

/** A selection of nothing, `width` × `height`. */
export function emptySelection(width, height) {
  const w = Math.max(1, Math.round(width));
  const h = Math.max(1, Math.round(height));
  return { width: w, height: h, data: new Uint8Array(w * h) };
}

/** Everything selected. */
export function selectAll(width, height) {
  const s = emptySelection(width, height);
  s.data.fill(255);
  return s;
}

/** How much of the span [a, b) a pixel [i, i + 1) covers, 0 to 1. */
const cover = (i, a, b) => clamp(Math.min(i + 1, b) - Math.max(i, a), 0, 1);

/** A rectangle `{ x, y, width, height }` in canvas pixels; fractional edges are partly selected. */
export function rectSelection(width, height, r) {
  const s = emptySelection(width, height);
  const x0 = Math.min(r.x, r.x + r.width);
  const x1 = Math.max(r.x, r.x + r.width);
  const y0 = Math.min(r.y, r.y + r.height);
  const y1 = Math.max(r.y, r.y + r.height);
  const ya = Math.max(0, Math.floor(y0));
  const yb = Math.min(s.height, Math.ceil(y1));
  const xa = Math.max(0, Math.floor(x0));
  const xb = Math.min(s.width, Math.ceil(x1));
  for (let y = ya; y < yb; y++) {
    const cy = cover(y, y0, y1);
    for (let x = xa; x < xb; x++) s.data[y * s.width + x] = Math.round(255 * cy * cover(x, x0, x1));
  }
  return s;
}

/** The ellipse inside a rectangle, its edge smoothed by sampling each edge pixel four by four. */
export function ellipseSelection(width, height, r) {
  const s = emptySelection(width, height);
  const x0 = Math.min(r.x, r.x + r.width);
  const y0 = Math.min(r.y, r.y + r.height);
  const rx = Math.abs(r.width) / 2;
  const ry = Math.abs(r.height) / 2;
  if (rx <= 0 || ry <= 0) return s;
  const cx = x0 + rx;
  const cy = y0 + ry;
  const inside = (x, y) => ((x - cx) / rx) ** 2 + ((y - cy) / ry) ** 2 <= 1;
  const ya = Math.max(0, Math.floor(y0));
  const yb = Math.min(s.height, Math.ceil(y0 + 2 * ry));
  const xa = Math.max(0, Math.floor(x0));
  const xb = Math.min(s.width, Math.ceil(x0 + 2 * rx));
  for (let y = ya; y < yb; y++) {
    for (let x = xa; x < xb; x++) {
      // The four corners decide most pixels; only an edge pixel is sampled finely.
      const c = inside(x, y) + inside(x + 1, y) + inside(x, y + 1) + inside(x + 1, y + 1);
      if (c === 4) { s.data[y * s.width + x] = 255; continue; }
      let n = 0;
      for (let j = 0; j < 4; j++) for (let i = 0; i < 4; i++) n += inside(x + (i + 0.5) / 4, y + (j + 0.5) / 4);
      if (n) s.data[y * s.width + x] = Math.round((255 * n) / 16);
    }
  }
  return s;
}

/**
 * A lasso: the shape a path of points `[[x, y], …]` closes, filled even-odd,
 * four sub-rows to a pixel with the spans' ends partly covered, so the edge
 * is smooth however the path runs.
 */
export function polygonSelection(width, height, points) {
  const s = emptySelection(width, height);
  if (!Array.isArray(points) || points.length < 3) return s;
  const W = s.width;
  const acc = new Float32Array(W);
  const ys = points.map((p) => p[1]);
  const ya = Math.max(0, Math.floor(Math.min(...ys)));
  const yb = Math.min(s.height, Math.ceil(Math.max(...ys)));
  const SUB = 4;
  for (let y = ya; y < yb; y++) {
    acc.fill(0);
    let any = false;
    for (let k = 0; k < SUB; k++) {
      const sy = y + (k + 0.5) / SUB;
      const xs = [];
      for (let i = 0; i < points.length; i++) {
        const [ax, ay] = points[i];
        const [bx, by] = points[(i + 1) % points.length];
        if ((ay <= sy && by > sy) || (by <= sy && ay > sy)) xs.push(ax + ((sy - ay) / (by - ay)) * (bx - ax));
      }
      xs.sort((a, b) => a - b);
      for (let i = 0; i + 1 < xs.length; i += 2) {
        const a = clamp(xs[i], 0, W);
        const b = clamp(xs[i + 1], 0, W);
        if (b <= a) continue;
        any = true;
        for (let x = Math.floor(a); x < Math.min(W, Math.ceil(b)); x++) acc[x] += cover(x, a, b) / SUB;
      }
    }
    if (!any) continue;
    const row = y * W;
    for (let x = 0; x < W; x++) if (acc[x] > 0) s.data[row + x] = Math.round(255 * Math.min(1, acc[x]));
  }
  return s;
}

/** How far apart two colours are: the largest difference of one channel, alpha included. */
export function colourDistance(data, i, c) {
  const dr = Math.abs(data[i] - c[0]);
  const dg = Math.abs(data[i + 1] - c[1]);
  const db = Math.abs(data[i + 2] - c[2]);
  const da = Math.abs(data[i + 3] - (c[3] ?? 255));
  return Math.max(dr, dg, db, da);
}

/**
 * The magic wand: the pixels of an RGBA picture within `tolerance` (0 to
 * 255, on every channel) of the one at (x, y). `contiguous` keeps only those
 * joined to it (four ways); otherwise every such pixel in the picture. This is
 * the bucket fill's region too.
 */
export function magicWand(rgba, width, height, x, y, { tolerance = 32, contiguous = true } = {}) {
  const s = emptySelection(width, height);
  const px = Math.floor(x);
  const py = Math.floor(y);
  if (px < 0 || py < 0 || px >= width || py >= height) return s;
  const i0 = (py * width + px) * 4;
  const seed = [rgba[i0], rgba[i0 + 1], rgba[i0 + 2], rgba[i0 + 3]];
  if (!contiguous) return colourSelection(rgba, width, height, seed, { tolerance });
  const tol = clamp(Math.round(tolerance), 0, 255);
  const seen = new Uint8Array(width * height);
  const stack = new Int32Array(width * height);
  let top = 0;
  stack[top++] = py * width + px;
  seen[py * width + px] = 1;
  while (top) {
    const p = stack[--top];
    s.data[p] = 255;
    const x0 = p % width;
    const y0 = (p - x0) / width;
    const visit = (q) => {
      if (seen[q]) return;
      seen[q] = 1;
      if (colourDistance(rgba, q * 4, seed) <= tol) stack[top++] = q;
    };
    if (x0 > 0) visit(p - 1);
    if (x0 < width - 1) visit(p + 1);
    if (y0 > 0) visit(p - width);
    if (y0 < height - 1) visit(p + width);
  }
  return s;
}

/** Select by colour: every pixel within `tolerance` of `colour` ([r, g, b, a?]). */
export function colourSelection(rgba, width, height, colour, { tolerance = 32 } = {}) {
  const s = emptySelection(width, height);
  const tol = clamp(Math.round(tolerance), 0, 255);
  for (let p = 0, n = width * height; p < n; p++) if (colourDistance(rgba, p * 4, colour) <= tol) s.data[p] = 255;
  return s;
}

/**
 * Two selections made one: `replace` takes `next`, `add` the most of both,
 * `subtract` what `next` leaves of `base`, `intersect` what both hold.
 */
export function combineSelection(base, next, mode = 'replace') {
  if (!base || mode === 'replace') return next;
  const out = emptySelection(base.width, base.height);
  const a = base.data;
  const b = next.data;
  for (let i = 0; i < a.length; i++) {
    const x = a[i];
    const y = b[i] || 0;
    out.data[i] = mode === 'add' ? Math.max(x, y) : mode === 'subtract' ? Math.round((x * (255 - y)) / 255) : Math.round((x * y) / 255);
  }
  return out;
}

/** The selection turned inside out. */
export function invertSelection(sel) {
  const out = emptySelection(sel.width, sel.height);
  for (let i = 0; i < sel.data.length; i++) out.data[i] = 255 - sel.data[i];
  return out;
}

/** Whether nothing at all is selected. */
export function isEmptySelection(sel) {
  if (!sel) return true;
  for (let i = 0; i < sel.data.length; i++) if (sel.data[i]) return false;
  return true;
}

/** The box round everything selected more than `threshold`: { x, y, width, height }, or null. */
export function selectionBounds(sel, threshold = 0) {
  if (!sel) return null;
  const { width: W, height: H, data } = sel;
  let x0 = W;
  let y0 = H;
  let x1 = -1;
  let y1 = -1;
  for (let y = 0; y < H; y++) {
    const row = y * W;
    for (let x = 0; x < W; x++) {
      if (data[row + x] > threshold) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        y1 = y;
      }
    }
  }
  return x1 < 0 ? null : { x: x0, y: y0, width: x1 - x0 + 1, height: y1 - y0 + 1 };
}

/** One pass of a box blur of radius `r` along rows (`horizontal`) or columns, on floats. */
function boxPass(src, dst, W, H, r, horizontal) {
  const n = horizontal ? W : H;
  const lines = horizontal ? H : W;
  const step = horizontal ? 1 : W;
  const span = 2 * r + 1;
  for (let l = 0; l < lines; l++) {
    const base = horizontal ? l * W : l;
    let sum = 0;
    // Edges held: the first and last values stand for what lies past them.
    for (let k = -r; k <= r; k++) sum += src[base + clamp(k, 0, n - 1) * step];
    for (let i = 0; i < n; i++) {
      dst[base + i * step] = sum / span;
      const out = clamp(i - r, 0, n - 1);
      const inn = clamp(i + r + 1, 0, n - 1);
      sum += src[base + inn * step] - src[base + out * step];
    }
  }
}

/**
 * Blur an alpha plane (a Float32Array or a Uint8Array) of W × H by about
 * `radius` pixels: three box passes each way, which is close to a Gaussian
 * whose spread is `radius` / 2. Answers a Float32Array.
 */
export function blurPlane(plane, W, H, radius) {
  let a = Float32Array.from(plane);
  const r = Math.max(0, radius);
  if (r < 0.5) return a;
  // Three boxes whose widths add up to a Gaussian of sigma = radius / 2.
  const sigma = r / 2;
  const ideal = Math.sqrt((12 * sigma * sigma) / 3 + 1);
  let wl = Math.floor(ideal);
  if (wl % 2 === 0) wl -= 1;
  const m = Math.round((12 * sigma * sigma - 3 * wl * wl - 12 * wl - 9) / (-4 * wl - 4));
  let b = new Float32Array(a.length);
  for (let i = 0; i < 3; i++) {
    const rr = Math.max(0, ((i < m ? wl : wl + 2) - 1) / 2);
    if (!rr) continue;
    boxPass(a, b, W, H, rr, true);
    boxPass(b, a, W, H, rr, false);
  }
  b = null;
  return a;
}

/** Feather: the selection's edge softened over about `radius` pixels either side. */
export function featherSelection(sel, radius) {
  if (!(radius > 0)) return sel;
  const out = emptySelection(sel.width, sel.height);
  const blurred = blurPlane(sel.data, sel.width, sel.height, radius);
  for (let i = 0; i < blurred.length; i++) out.data[i] = Math.round(clamp(blurred[i], 0, 255));
  return out;
}

/**
 * Distance, in pixels, from every pixel to the nearest pixel where `on` is
 * true, exactly (the squared Euclidean transform of Felzenszwalb and
 * Huttenlocher, a row pass then a column pass). Answers a Float32Array of
 * squared distances.
 */
export function distanceSquared(on, W, H) {
  const INF = 1e20;
  const f = new Float32Array(Math.max(W, H));
  const d = new Float32Array(Math.max(W, H));
  const v = new Int32Array(Math.max(W, H));
  const z = new Float32Array(Math.max(W, H) + 1);
  const grid = new Float32Array(W * H);
  for (let i = 0; i < W * H; i++) grid[i] = on(i) ? 0 : INF;
  const pass = (n, get, set) => {
    for (let q = 0; q < n; q++) f[q] = get(q);
    let k = 0;
    v[0] = 0;
    z[0] = -INF;
    z[1] = INF;
    for (let q = 1; q < n; q++) {
      let s = ((f[q] + q * q) - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
      while (s <= z[k]) {
        k--;
        s = ((f[q] + q * q) - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
      }
      k++;
      v[k] = q;
      z[k] = s;
      z[k + 1] = INF;
    }
    k = 0;
    for (let q = 0; q < n; q++) {
      while (z[k + 1] < q) k++;
      d[q] = (q - v[k]) * (q - v[k]) + f[v[k]];
    }
    for (let q = 0; q < n; q++) set(q, d[q]);
  };
  for (let x = 0; x < W; x++) pass(H, (y) => grid[y * W + x], (y, val) => { grid[y * W + x] = val; });
  for (let y = 0; y < H; y++) pass(W, (x) => grid[y * W + x], (x, val) => { grid[y * W + x] = val; });
  return grid;
}

/** Grow: the selection widened by `px` pixels all round, its corners rounded. */
export function growSelection(sel, px) {
  const r = Math.max(0, Number(px) || 0);
  if (!r) return sel;
  const { width: W, height: H } = sel;
  const dist = distanceSquared((i) => sel.data[i] >= 128, W, H);
  const out = emptySelection(W, H);
  for (let i = 0; i < dist.length; i++) {
    if (sel.data[i] >= 128) { out.data[i] = 255; continue; }
    const d = Math.sqrt(dist[i]);
    // A pixel's worth of soft edge at the new boundary.
    out.data[i] = Math.round(255 * clamp(r + 1 - d, 0, 1));
  }
  return out;
}

/** Shrink: the selection narrowed by `px` pixels all round. */
export function shrinkSelection(sel, px) {
  const r = Math.max(0, Number(px) || 0);
  if (!r) return sel;
  const { width: W, height: H } = sel;
  const dist = distanceSquared((i) => sel.data[i] < 128, W, H);
  const out = emptySelection(W, H);
  for (let i = 0; i < dist.length; i++) {
    if (sel.data[i] < 128) continue;
    const d = Math.sqrt(dist[i]);
    out.data[i] = Math.round(sel.data[i] * clamp(d - r, 0, 1));
  }
  return out;
}

/**
 * The outline marching ants run along: every edge between a pixel that is
 * selected (128 or more) and one that is not, joined into runs, as a flat
 * list of segments [x0, y0, x1, y1, …] in canvas pixels.
 */
export function selectionOutline(sel, threshold = 128) {
  const { width: W, height: H, data } = sel;
  const inside = (x, y) => x >= 0 && y >= 0 && x < W && y < H && data[y * W + x] >= threshold;
  const out = [];
  // Horizontal edges: between row y - 1 and row y.
  for (let y = 0; y <= H; y++) {
    let start = -1;
    for (let x = 0; x <= W; x++) {
      const edge = x < W && inside(x, y - 1) !== inside(x, y);
      if (edge && start < 0) start = x;
      if (!edge && start >= 0) { out.push(start, y, x, y); start = -1; }
    }
  }
  // Vertical edges: between column x - 1 and column x.
  for (let x = 0; x <= W; x++) {
    let start = -1;
    for (let y = 0; y <= H; y++) {
      const edge = y < H && inside(x - 1, y) !== inside(x, y);
      if (edge && start < 0) start = y;
      if (!edge && start >= 0) { out.push(x, start, x, y); start = -1; }
    }
  }
  return out;
}

/**
 * A mask laid over an RGBA picture: every pixel's alpha multiplied by the
 * mask's (or by what the mask leaves out, `invert`). The painted mask of a
 * layer, made part of its pixels by Apply, and Delete inside a selection
 * (`invert`: what is selected goes) both come here.
 */
export function applyMask(rgba, mask, { invert = false } = {}) {
  const out = new Uint8ClampedArray(rgba);
  for (let p = 0, n = mask.length; p < n; p++) {
    const m = invert ? 255 - mask[p] : mask[p];
    out[p * 4 + 3] = Math.round((rgba[p * 4 + 3] * m) / 255);
  }
  return out;
}

/**
 * What a painted mask lets through at one pixel, 0 to 1: its own alpha
 * where it lies, `outside` (0 hide or 1 show) beyond its edges, turned over
 * by `invert`. A disabled mask lets everything through.
 */
export function maskCoverage(maskAlpha, { inside = true, outside = 1, invert = false, off = false } = {}) {
  if (off) return 1;
  const a = inside ? maskAlpha / 255 : outside;
  return invert ? 1 - a : a;
}

/*
 * A selection kept in a file (the Image project, docs/STUDIO.md 7.2): its
 * bytes run by run, each run a count (a variable-length number) and the
 * value it repeats. A selection is mostly long runs of 0 and 255 with a
 * feathered edge between, so a canvas-sized plane packs to a few kilobytes,
 * exactly: unpacked, it is the plane that was packed, value for value.
 *
 *   "RSEL", version 1, width and height (four bytes each, big-endian), runs
 */
const SEL_MAGIC = [0x52, 0x53, 0x45, 0x4c];

/** A selection as bytes for a file. */
export function packSelection(sel) {
  const { width, height, data } = sel;
  const out = [...SEL_MAGIC, 1, (width >>> 24) & 0xff, (width >>> 16) & 0xff, (width >>> 8) & 0xff, width & 0xff, (height >>> 24) & 0xff, (height >>> 16) & 0xff, (height >>> 8) & 0xff, height & 0xff];
  let i = 0;
  while (i < data.length) {
    const v = data[i];
    let n = 1;
    while (i + n < data.length && data[i + n] === v) n++;
    let c = n;
    while (c >= 0x80) {
      out.push((c & 0x7f) | 0x80);
      c = Math.floor(c / 128);
    }
    out.push(c, v);
    i += n;
  }
  return Uint8Array.from(out);
}

/** Bytes `packSelection` wrote, as the selection again; null for anything else or anything damaged. */
export function unpackSelection(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []);
  if (b.length < 13 || SEL_MAGIC.some((m, k) => b[k] !== m) || b[4] !== 1) return null;
  const width = ((b[5] << 24) | (b[6] << 16) | (b[7] << 8) | b[8]) >>> 0;
  const height = ((b[9] << 24) | (b[10] << 16) | (b[11] << 8) | b[12]) >>> 0;
  if (!width || !height || width * height > 400e6) return null;
  const data = new Uint8Array(width * height);
  let at = 13;
  let i = 0;
  while (at < b.length && i < data.length) {
    let n = 0;
    let scale = 1;
    for (;;) {
      const c = b[at++];
      if (c === undefined) return null;
      n += (c & 0x7f) * scale;
      if (!(c & 0x80)) break;
      scale *= 128;
    }
    const v = b[at++];
    if (v === undefined || i + n > data.length) return null;
    data.fill(v, i, i + n);
    i += n;
  }
  // Every pixel filled, and nothing left over: anything else is not what was packed.
  if (i !== data.length || at !== b.length) return null;
  return { width, height, data };
}
