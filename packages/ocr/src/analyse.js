// Page analysis of our own, before and after the recogniser: how far a scan
// is turned, which pixels are ink, how tall its text is, and where its
// ruling lines run. Classical image processing, deterministic and proved by
// the engine tests, not a model: on scans and rendered pages it is as good
// as one and easier to check.
//
// Every function takes the grey or binary pictures of pixels.js:
// { width, height, data } with one byte a pixel; in a binary picture 0 is
// ink and 255 paper.

/** The threshold that best splits the picture's grey levels in two (Otsu, 1979). */
export function otsu({ data }) {
  const hist = new Float64Array(256);
  for (let i = 0; i < data.length; i++) hist[data[i]]++;
  const total = data.length;
  let sum = 0;
  for (let i = 0; i < 256; i++) sum += i * hist[i];
  let sumB = 0, wB = 0, best = 0, threshold = 127;
  for (let t = 0; t < 256; t++) {
    wB += hist[t];
    if (!wB) continue;
    const wF = total - wB;
    if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sum - sumB) / wF;
    const between = wB * wF * (mB - mF) * (mB - mF);
    if (between > best) { best = between; threshold = t; }
  }
  return threshold;
}

/** Sums over every rectangle in constant time: the integral of values and of their squares. */
function integrals({ width, height, data }) {
  const w = width + 1;
  const s = new Float64Array(w * (height + 1));
  const q = new Float64Array(w * (height + 1));
  for (let y = 0; y < height; y++) {
    let rs = 0, rq = 0;
    for (let x = 0; x < width; x++) {
      const v = data[y * width + x];
      rs += v; rq += v * v;
      s[(y + 1) * w + x + 1] = s[y * w + x + 1] + rs;
      q[(y + 1) * w + x + 1] = q[y * w + x + 1] + rq;
    }
  }
  return { s, q, w };
}

/**
 * Ink and paper. 'otsu' splits the whole page at one level; 'sauvola' (2000)
 * judges each pixel against the mean and spread of its neighbourhood, so
 * a page lit unevenly, or with a shadow down its fold, keeps its text.
 * 'auto' picks Sauvola when the paper's brightness varies across the page.
 */
export function binarise(grey, { method = 'auto', window = 0, k = 0.2 } = {}) {
  const { width, height, data } = grey;
  const out = new Uint8Array(width * height);
  const use = method === 'auto' ? (unevenPaper(grey) ? 'sauvola' : 'otsu') : method;
  if (use === 'otsu') {
    const t = otsu(grey);
    for (let i = 0; i < data.length; i++) out[i] = data[i] <= t ? 0 : 255;
    return { width, height, data: out, method: 'otsu', threshold: t };
  }
  const r = Math.max(7, (window || Math.round(Math.min(width, height) / 40)) >> 1);
  const { s, q, w } = integrals(grey);
  const R = 128;
  for (let y = 0; y < height; y++) {
    const y0 = Math.max(0, y - r), y1 = Math.min(height, y + r + 1);
    for (let x = 0; x < width; x++) {
      const x0 = Math.max(0, x - r), x1 = Math.min(width, x + r + 1);
      const n = (x1 - x0) * (y1 - y0);
      const sum = s[y1 * w + x1] - s[y0 * w + x1] - s[y1 * w + x0] + s[y0 * w + x0];
      const sq = q[y1 * w + x1] - q[y0 * w + x1] - q[y1 * w + x0] + q[y0 * w + x0];
      const mean = sum / n;
      const sd = Math.sqrt(Math.max(0, sq / n - mean * mean));
      out[y * width + x] = data[y * width + x] <= mean * (1 + k * (sd / R - 1)) ? 0 : 255;
    }
  }
  return { width, height, data: out, method: 'sauvola' };
}

/** Whether the paper is brighter in some parts of the page than others by more than a few levels. */
export function unevenPaper({ width, height, data }) {
  const cells = 6;
  const bright = [];
  for (let cy = 0; cy < cells; cy++) {
    for (let cx = 0; cx < cells; cx++) {
      const hist = new Uint32Array(256);
      let n = 0;
      for (let y = Math.floor((cy * height) / cells); y < Math.floor(((cy + 1) * height) / cells); y += 2) {
        for (let x = Math.floor((cx * width) / cells); x < Math.floor(((cx + 1) * width) / cells); x += 2) { hist[data[y * width + x]]++; n++; }
      }
      // The paper's level in this cell: the 90th percentile, above the ink.
      let acc = 0, level = 255;
      for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc >= n * 0.9) { level = v; break; } }
      bright.push(level);
    }
  }
  return Math.max(...bright) - Math.min(...bright) > 40;
}

/** The ink pixels of a binary picture, every `step`th in each direction. */
function inkPoints({ width, height, data }, step) {
  const xs = [], ys = [];
  for (let y = 0; y < height; y += step) for (let x = 0; x < width; x += step) if (data[y * width + x] === 0) { xs.push(x); ys.push(y); }
  return { xs, ys };
}

/**
 * How far the lines of text are turned, in degrees, counter-clockwise
 * positive (the way to turn the page back is the negative). The rows of
 * ink are projected at each candidate angle; at the right one the text
 * lines fall into sharp peaks and the gaps between them into empty rows,
 * which maximises the sum of squared differences between neighbouring
 * rows. Coarse steps over plus or minus `range`, then fine ones.
 */
export function estimateSkew(binary, { range = 8 } = {}) {
  const { width, height } = binary;
  const step = Math.max(1, Math.round(Math.sqrt((width * height) / 400000)));
  const { xs, ys } = inkPoints(binary, step);
  if (xs.length < 50) return 0;
  const cx = width / 2;
  const score = (deg) => {
    const t = Math.tan((deg * Math.PI) / 180);
    // Bins as tall as the sampling step: finer ones would alternate full and empty at 0 degrees, and win there.
    const bins = new Float64Array(Math.ceil((height + width) / step) + 2);
    const off = Math.ceil((width * Math.abs(t)) / step) + 1;
    for (let i = 0; i < xs.length; i++) {
      const b = Math.round((ys[i] + (xs[i] - cx) * t) / step) + off;
      if (b >= 0 && b < bins.length) bins[b]++;
    }
    let s = 0;
    for (let i = 1; i < bins.length; i++) { const d = bins[i] - bins[i - 1]; s += d * d; }
    return s;
  };
  const search = (from, to, by) => {
    let best = 0, bestScore = -1;
    for (let a = from; a <= to + 1e-9; a += by) {
      const s = score(a);
      if (s > bestScore) { bestScore = s; best = a; }
    }
    return best;
  };
  const coarse = search(-range, range, 0.5);
  const fine = search(coarse - 0.5, coarse + 0.5, 0.05);
  // A line turned counter-clockwise by a rises by tan(a) a pixel to the right; the shear that levels it is the same angle.
  return Math.round(fine * 100) / 100;
}

/** Bilinear sample of a grey picture, paper outside it. */
function sampler({ width, height, data }, fill) {
  return (x, y) => {
    const x0 = Math.floor(x), y0 = Math.floor(y);
    if (x0 < -1 || y0 < -1 || x0 >= width || y0 >= height) return fill;
    const fx = x - x0, fy = y - y0;
    const at = (xx, yy) => (xx < 0 || yy < 0 || xx >= width || yy >= height ? fill : data[yy * width + xx]);
    return (at(x0, y0) * (1 - fx) + at(x0 + 1, y0) * fx) * (1 - fy) + (at(x0, y0 + 1) * (1 - fx) + at(x0 + 1, y0 + 1) * fx) * fy;
  };
}

/** The picture turned by `degrees` counter-clockwise about its centre, the same size, paper where it uncovers. */
export function rotate(grey, degrees, { fill = 255 } = {}) {
  const { width, height } = grey;
  const out = new Uint8Array(width * height);
  const a = (degrees * Math.PI) / 180;
  const cos = Math.cos(a), sin = Math.sin(a);
  const cx = (width - 1) / 2, cy = (height - 1) / 2;
  const sample = sampler(grey, fill);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      // The source of each output pixel: turned back. y grows downwards, so counter-clockwise on screen is this sign.
      const dx = x - cx, dy = y - cy;
      out[y * width + x] = sample(cx + dx * cos - dy * sin, cy + dx * sin + dy * cos);
    }
  }
  return { width, height, data: out };
}

/** Where a point of the turned picture was in the original. */
export function unrotatePoint(x, y, degrees, width, height) {
  const a = (degrees * Math.PI) / 180;
  const cx = (width - 1) / 2, cy = (height - 1) / 2;
  const dx = x - cx, dy = y - cy;
  return { x: cx + dx * Math.cos(a) - dy * Math.sin(a), y: cy + dx * Math.sin(a) + dy * Math.cos(a) };
}

/** The picture scaled by `factor`: bilinear when growing, the mean of each area when shrinking. */
export function scale(grey, factor) {
  const { width, height, data } = grey;
  const w = Math.max(1, Math.round(width * factor)), h = Math.max(1, Math.round(height * factor));
  const out = new Uint8Array(w * h);
  if (factor >= 1) {
    const sample = sampler(grey, 255);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) out[y * w + x] = sample((x + 0.5) / factor - 0.5, (y + 0.5) / factor - 0.5);
  } else {
    for (let y = 0; y < h; y++) {
      const y0 = Math.floor(y / factor), y1 = Math.max(y0 + 1, Math.min(height, Math.floor((y + 1) / factor)));
      for (let x = 0; x < w; x++) {
        const x0 = Math.floor(x / factor), x1 = Math.max(x0 + 1, Math.min(width, Math.floor((x + 1) / factor)));
        let s = 0;
        for (let yy = y0; yy < y1; yy++) for (let xx = x0; xx < x1; xx++) s += data[yy * width + xx];
        out[y * w + x] = s / ((y1 - y0) * (x1 - x0));
      }
    }
  }
  return { width: w, height: h, data: out };
}

/** A part of a grey picture, with `pad` pixels of paper round it (the engine reads a word best with a margin). */
export function cropPadded(grey, x, y, w, h, pad = 12, fill = 255) {
  const x0 = Math.max(0, Math.floor(x)), y0 = Math.max(0, Math.floor(y));
  const x1 = Math.min(grey.width, Math.ceil(x + w)), y1 = Math.min(grey.height, Math.ceil(y + h));
  const cw = Math.max(1, x1 - x0), ch = Math.max(1, y1 - y0);
  const out = { width: cw + pad * 2, height: ch + pad * 2, data: new Uint8Array((cw + pad * 2) * (ch + pad * 2)).fill(fill) };
  for (let yy = 0; yy < ch; yy++) out.data.set(grey.data.subarray((y0 + yy) * grey.width + x0, (y0 + yy) * grey.width + x0 + cw), (yy + pad) * out.width + pad);
  return out;
}

/** How many ink pixels a binary picture has inside a rectangle. */
export function inkIn(binary, x, y, w, h) {
  let n = 0;
  for (let yy = Math.max(0, Math.floor(y)); yy < Math.min(binary.height, Math.ceil(y + h)); yy++) {
    for (let xx = Math.max(0, Math.floor(x)); xx < Math.min(binary.width, Math.ceil(x + w)); xx++) if (binary.data[yy * binary.width + xx] === 0) n++;
  }
  return n;
}

/**
 * The connected blobs of ink (eight-connected), each with its box and
 * pixel count; runs are labelled row by row and joined with union-find.
 */
export function components({ width, height, data }) {
  const parent = [];
  const find = (a) => { while (parent[a] !== a) { parent[a] = parent[parent[a]]; a = parent[a]; } return a; };
  const union = (a, b) => { a = find(a); b = find(b); if (a !== b) parent[Math.max(a, b)] = Math.min(a, b); };
  let prev = [];
  const runs = [];
  for (let y = 0; y < height; y++) {
    const row = [];
    let x = 0;
    while (x < width) {
      if (data[y * width + x] !== 0) { x++; continue; }
      const start = x;
      while (x < width && data[y * width + x] === 0) x++;
      const id = parent.length;
      parent.push(id);
      const run = { y, x0: start, x1: x - 1, id };
      for (const p of prev) if (p.x0 <= run.x1 + 1 && p.x1 >= run.x0 - 1) union(id, p.id);
      row.push(run);
      runs.push(run);
    }
    prev = row;
  }
  const blobs = new Map();
  for (const r of runs) {
    const root = find(r.id);
    let b = blobs.get(root);
    if (!b) blobs.set(root, (b = { x0: r.x0, y0: r.y, x1: r.x1, y1: r.y, pixels: 0 }));
    b.x0 = Math.min(b.x0, r.x0); b.x1 = Math.max(b.x1, r.x1);
    b.y0 = Math.min(b.y0, r.y); b.y1 = Math.max(b.y1, r.y);
    b.pixels += r.x1 - r.x0 + 1;
  }
  return [...blobs.values()].map((b) => ({ x: b.x0, y: b.y0, width: b.x1 - b.x0 + 1, height: b.y1 - b.y0 + 1, pixels: b.pixels }));
}

/**
 * The height of the page's text in pixels: the median height of the blobs
 * that look like letters (not specks, not rules, not pictures), each
 * counted by its ink, so the dots of Arabic script and the parts of a Han
 * character, small and many, do not pull it down. 0 when there is no text
 * to measure.
 */
export function textHeight(binary) {
  const blobs = components(binary).filter((b) => b.height >= 4 && b.width >= 2 && b.height < binary.height / 4 && b.width < binary.width / 4 && b.width / b.height < 6 && b.height / b.width < 8);
  if (blobs.length < 5) return 0;
  blobs.sort((a, b) => a.height - b.height);
  const half = blobs.reduce((n, b) => n + b.pixels, 0) / 2;
  let acc = 0;
  for (const b of blobs) if ((acc += b.pixels) >= half) return b.height;
  return blobs[blobs.length - 1].height;
}

/**
 * The ruling lines of a page: runs of ink longer than `min` pixels along a
 * row (horizontal) or down a column (vertical), the runs of neighbouring
 * rows merged into one line with its thickness.
 * @returns {{ horizontal: { y, x0, x1, thickness }[], vertical: { x, y0, y1, thickness }[] }}
 */
export function rulings(binary, { min = 0 } = {}) {
  const { width, height, data } = binary;
  const minH = min || Math.max(40, Math.round(width / 12));
  const minV = min || Math.max(30, Math.round(height / 30));
  const along = (count, length, at, least) => {
    // Every long run in each lane, then runs in touching lanes that overlap merged.
    const lines = [];
    let open = [];
    for (let lane = 0; lane < count; lane++) {
      const here = [];
      // The lane's runs of ink, a break of under three pixels (a scan's dropout) bridged.
      let i = 0, last = null;
      while (i < length) {
        if (at(lane, i) !== 0) { i++; continue; }
        const s = i;
        while (i < length && at(lane, i) === 0) i++;
        if (last && s - last.b - 1 < 3) last.b = i - 1;
        else here.push((last = { a: s, b: i - 1, from: lane, to: lane }));
      }
      for (let r = here.length - 1; r >= 0; r--) if (here[r].b - here[r].a + 1 < least) here.splice(r, 1);
      const next = [];
      for (const run of here) {
        const join = open.find((o) => o.to === lane - 1 && o.a <= run.b && o.b >= run.a);
        if (join) { join.to = lane; join.a = Math.min(join.a, run.a); join.b = Math.max(join.b, run.b); next.push(join); }
        else { lines.push(run); next.push(run); }
      }
      open = next;
    }
    return lines;
  };
  const horizontal = along(height, width, (y, x) => data[y * width + x], minH)
    .filter((l) => l.to - l.from + 1 <= Math.max(6, Math.round(height / 100)))
    .map((l) => ({ y: (l.from + l.to) / 2, x0: l.a, x1: l.b, thickness: l.to - l.from + 1 }));
  const vertical = along(width, height, (x, y) => data[y * width + x], minV)
    .filter((l) => l.to - l.from + 1 <= Math.max(6, Math.round(width / 100)))
    .map((l) => ({ x: (l.from + l.to) / 2, y0: l.a, y1: l.b, thickness: l.to - l.from + 1 }));
  return { horizontal, vertical };
}
