// Work on the pixels themselves: what a canvas filter cannot say.
//
// Levels, curves, white balance, exposure, highlights and shadows, sharpen,
// noise reduction, vignette and red-eye, each a pure function over an
// ImageData-shaped `{ data, width, height }` (RGBA, 8 bits a channel). No
// DOM: the same functions run in a test in Node, in the window, and in the
// image worker that keeps a large picture off the window's thread.
//
// The tone operations are point operations, so they are folded into lookup
// tables once per call and the pixel loop is a few array reads. The order
// follows a darkroom: white balance on the light as it came in, then
// exposure, then levels, then highlights and shadows on the luminance, then
// the curve on the finished tones.

/** The settings, in the units the window's sliders show. */
export const DEFAULT_PIXELS = {
  exposure: 0,      // stops, -3 to 3
  highlights: 0,    // -100 to 100: negative brings bright parts down
  shadows: 0,       // -100 to 100: positive lifts the dark parts
  temperature: 0,   // -100 (cooler, bluer) to 100 (warmer, yellower)
  tint: 0,          // -100 (greener) to 100 (more magenta)
  levels: null,     // { inBlack, inWhite, gamma, outBlack, outWhite }
  curve: null,      // { rgb, r, g, b }: each [[in, out], ...] in 0..255
  sharpen: 0,       // 0 to 100
  denoise: 0,       // 0 to 100
  vignette: 0,      // -100 (lighter edges) to 100 (darker edges)
};

export const NEUTRAL_LEVELS = { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 };

const clamp = (v, lo, hi) => (v < lo ? lo : v > hi ? hi : v);
const byte = (v) => (v < 0 ? 0 : v > 255 ? 255 : Math.round(v));
const num = (v, fallback = 0) => (Number.isFinite(Number(v)) ? Number(v) : fallback);

/** True when the settings would change no pixel. */
export function isNeutral(settings = {}) {
  const s = { ...DEFAULT_PIXELS, ...settings };
  if (s.exposure || s.highlights || s.shadows || s.temperature || s.tint || s.sharpen || s.denoise || s.vignette) return false;
  const l = { ...NEUTRAL_LEVELS, ...(s.levels || {}) };
  if (l.inBlack !== 0 || l.inWhite !== 255 || l.gamma !== 1 || l.outBlack !== 0 || l.outWhite !== 255) return false;
  if (s.curve && Object.values(s.curve).some((pts) => Array.isArray(pts) && pts.some(([x, y]) => x !== y))) return false;
  return true;
}

// ── lookup tables ───────────────────────────────────────────────────────────

/** Levels as a table: input black and white points, gamma, output black and white. */
export function levelsTable(levels = {}) {
  const l = { ...NEUTRAL_LEVELS, ...levels };
  const inBlack = clamp(num(l.inBlack), 0, 254);
  const inWhite = clamp(num(l.inWhite, 255), inBlack + 1, 255);
  const gamma = clamp(num(l.gamma, 1), 0.1, 10);
  const outBlack = clamp(num(l.outBlack), 0, 255);
  const outWhite = clamp(num(l.outWhite, 255), 0, 255);
  const out = new Uint8ClampedArray(256);
  for (let i = 0; i < 256; i++) {
    let v = clamp((i - inBlack) / (inWhite - inBlack), 0, 1);
    if (gamma !== 1) v = Math.pow(v, 1 / gamma);
    out[i] = byte(outBlack + v * (outWhite - outBlack));
  }
  return out;
}

/**
 * A curve through points as a table, by a monotone cubic (Fritsch and
 * Carlson): smooth like the curve tools people know, and never overshooting
 * between two points, so a curve cannot invent a bright rim the person did
 * not draw. Points are [[in, out], ...] in 0..255; fewer than two is the
 * straight line.
 */
export function curveTable(points) {
  const out = new Uint8ClampedArray(256);
  if (!Array.isArray(points) || points.length < 2) {
    for (let i = 0; i < 256; i++) out[i] = i;
    return out;
  }
  const pts = points
    .map((p) => [clamp(num(p[0]), 0, 255), clamp(num(p[1]), 0, 255)])
    .sort((p, q) => p[0] - q[0])
    .filter((p, i, all) => i === 0 || p[0] > all[i - 1][0]);
  const n = pts.length;
  if (n < 2) {
    out.fill(byte(pts[0][1]));
    return out;
  }
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const d = [];
  const m = [];
  for (let i = 0; i < n - 1; i++) d.push((ys[i + 1] - ys[i]) / (xs[i + 1] - xs[i]));
  m[0] = d[0];
  m[n - 1] = d[n - 2];
  for (let i = 1; i < n - 1; i++) m[i] = d[i - 1] * d[i] <= 0 ? 0 : (d[i - 1] + d[i]) / 2;
  for (let i = 0; i < n - 1; i++) {
    if (d[i] === 0) { m[i] = 0; m[i + 1] = 0; continue; }
    const a = m[i] / d[i];
    const b = m[i + 1] / d[i];
    const s = a * a + b * b;
    if (s > 9) {
      const t = 3 / Math.sqrt(s);
      m[i] = t * a * d[i];
      m[i + 1] = t * b * d[i];
    }
  }
  let seg = 0;
  for (let x = 0; x < 256; x++) {
    if (x <= xs[0]) { out[x] = byte(ys[0]); continue; }
    if (x >= xs[n - 1]) { out[x] = byte(ys[n - 1]); continue; }
    while (seg < n - 2 && x > xs[seg + 1]) seg++;
    const h = xs[seg + 1] - xs[seg];
    const t = (x - xs[seg]) / h;
    const t2 = t * t;
    const t3 = t2 * t;
    const v = (2 * t3 - 3 * t2 + 1) * ys[seg] + (t3 - 2 * t2 + t) * h * m[seg] + (-2 * t3 + 3 * t2) * ys[seg + 1] + (t3 - t2) * h * m[seg + 1];
    out[x] = byte(v);
  }
  return out;
}

/** Curves offered by name, as a curve dialog's presets are. */
export const CURVE_PRESETS = {
  linear: [[0, 0], [255, 255]],
  lighter: [[0, 0], [128, 166], [255, 255]],
  darker: [[0, 0], [128, 92], [255, 255]],
  contrast: [[0, 0], [64, 46], [192, 210], [255, 255]],
  softer: [[0, 0], [64, 80], [192, 176], [255, 255]],
};

/**
 * The three channel tables for white balance, exposure and levels, before
 * highlights and shadows, and the three for the curve after them.
 */
export function toneTables(settings = {}) {
  const s = { ...DEFAULT_PIXELS, ...settings };
  // White balance as a gain per channel. Temperature moves red against
  // blue, tint green against magenta, the two axes a camera offers; each
  // pair is balanced so the picture's brightness stays about where it was.
  const temp = clamp(num(s.temperature) / 100, -1, 1);
  const tint = clamp(num(s.tint) / 100, -1, 1);
  const gains = [1 + temp * 0.28 + tint * 0.08, 1 - tint * 0.2, 1 - temp * 0.28 + tint * 0.08];
  const exposure = Math.pow(2, clamp(num(s.exposure), -5, 5));
  const levels = levelsTable(s.levels || {});
  const curve = s.curve || {};
  const master = curveTable(curve.rgb);
  const first = [];
  const last = [];
  for (let c = 0; c < 3; c++) {
    const pre = new Uint8ClampedArray(256);
    for (let i = 0; i < 256; i++) {
      // Exposure in linear light, as a camera's would, then back to the screen's curve.
      const lin = Math.pow((i * gains[c]) / 255, 2.2) * exposure;
      const v = 255 * Math.pow(clamp(lin, 0, 1), 1 / 2.2);
      pre[i] = levels[byte(v)];
    }
    first.push(pre);
    const own = curveTable(curve[['r', 'g', 'b'][c]]);
    const post = new Uint8ClampedArray(256);
    for (let i = 0; i < 256; i++) post[i] = own[master[i]];
    last.push(post);
  }
  return { first, last };
}

/**
 * Highlights and shadows as a gain on each luminance: shadows lift (or sink)
 * the dark third, highlights the bright third, with black and white held
 * where they are and each colour kept by scaling its channels together.
 */
export function lightTable(highlights = 0, shadows = 0) {
  const h = clamp(num(highlights) / 100, -1, 1);
  const s = clamp(num(shadows) / 100, -1, 1);
  const gain = new Float32Array(256);
  for (let i = 0; i < 256; i++) {
    const l = i / 255;
    const shadowWeight = 1.5 * l * (1 - l) * (1 - l);
    const highlightWeight = 1.5 * l * l * (1 - l);
    const next = clamp(l + s * shadowWeight * 1.6 + h * highlightWeight * 1.6, 0, 1);
    gain[i] = l > 0 ? next / l : 1;
  }
  return gain;
}

// ── the passes ──────────────────────────────────────────────────────────────

/** Every point operation at once: tables, then highlights and shadows, then the curve. */
export function applyTone(img, settings = {}) {
  const { data } = img;
  const { first, last } = toneTables(settings);
  const [r0, g0, b0] = first;
  const [r1, g1, b1] = last;
  const light = settings.highlights || settings.shadows ? lightTable(settings.highlights, settings.shadows) : null;
  for (let i = 0; i < data.length; i += 4) {
    let r = r0[data[i]];
    let g = g0[data[i + 1]];
    let b = b0[data[i + 2]];
    if (light) {
      const k = light[(0.2126 * r + 0.7152 * g + 0.0722 * b + 0.5) | 0];
      r = byte(r * k);
      g = byte(g * k);
      b = byte(b * k);
    }
    data[i] = r1[r];
    data[i + 1] = g1[g];
    data[i + 2] = b1[b];
  }
  return img;
}

/** 255 less each channel; alpha kept. */
export function invertPixels(img) {
  const { data } = img;
  for (let i = 0; i < data.length; i += 4) {
    data[i] = 255 - data[i];
    data[i + 1] = 255 - data[i + 1];
    data[i + 2] = 255 - data[i + 2];
  }
  return img;
}

/** A box blur of radius r, horizontal then vertical, into a new buffer. */
function boxBlur(src, w, h, r) {
  const tmp = new Float32Array(w * h * 3);
  const out = new Float32Array(w * h * 3);
  const span = 2 * r + 1;
  for (let y = 0; y < h; y++) {
    for (let c = 0; c < 3; c++) {
      let sum = 0;
      for (let k = -r; k <= r; k++) sum += src[(y * w + clamp(k, 0, w - 1)) * 4 + c];
      for (let x = 0; x < w; x++) {
        tmp[(y * w + x) * 3 + c] = sum / span;
        const add = clamp(x + r + 1, 0, w - 1);
        const drop = clamp(x - r, 0, w - 1);
        sum += src[(y * w + add) * 4 + c] - src[(y * w + drop) * 4 + c];
      }
    }
  }
  for (let x = 0; x < w; x++) {
    for (let c = 0; c < 3; c++) {
      let sum = 0;
      for (let k = -r; k <= r; k++) sum += tmp[(clamp(k, 0, h - 1) * w + x) * 3 + c];
      for (let y = 0; y < h; y++) {
        out[(y * w + x) * 3 + c] = sum / span;
        const add = clamp(y + r + 1, 0, h - 1);
        const drop = clamp(y - r, 0, h - 1);
        sum += tmp[(add * w + x) * 3 + c] - tmp[(drop * w + x) * 3 + c];
      }
    }
  }
  return out;
}

/**
 * Sharpen by an unsharp mask: the picture plus `amount` of its difference
 * from a blurred copy. `radius` in pixels (1 for a small picture, more for a
 * large one, so the look keeps at every size); a small threshold leaves
 * flat parts, where the difference is noise, alone.
 */
export function sharpen(img, amount, { radius = 1, threshold = 2 } = {}) {
  const k = clamp(num(amount) / 100, 0, 1) * 1.6;
  if (!k) return img;
  const { data, width: w, height: h } = img;
  const r = Math.max(1, Math.round(radius));
  const blurred = boxBlur(data, w, h, r);
  for (let p = 0, q = 0; p < data.length; p += 4, q += 3) {
    for (let c = 0; c < 3; c++) {
      const diff = data[p + c] - blurred[q + c];
      if (Math.abs(diff) > threshold) data[p + c] = byte(data[p + c] + diff * k);
    }
  }
  return img;
}

/**
 * Noise reduction that keeps edges: each pixel becomes the average of its
 * neighbours weighted by how alike they are (a small bilateral filter), so
 * grain in a flat sky is smoothed and the line of a roof is not. Stronger
 * settings take a wider range and a second pass.
 */
export function denoise(img, strength) {
  const s = clamp(num(strength) / 100, 0, 1);
  if (!s) return img;
  const passes = s > 0.55 ? 2 : 1;
  const sigma = 8 + 36 * s;
  const weight = new Float32Array(766);
  for (let d = 0; d < weight.length; d++) weight[d] = Math.exp(-(d * d) / (2 * sigma * sigma * 3));
  const { width: w, height: h } = img;
  let src = img.data;
  for (let pass = 0; pass < passes; pass++) {
    const out = new Uint8ClampedArray(src.length);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4;
        const r = src[i];
        const g = src[i + 1];
        const b = src[i + 2];
        let sr = 0;
        let sg = 0;
        let sb = 0;
        let sw = 0;
        for (let dy = -1; dy <= 1; dy++) {
          const yy = y + dy;
          if (yy < 0 || yy >= h) continue;
          for (let dx = -1; dx <= 1; dx++) {
            const xx = x + dx;
            if (xx < 0 || xx >= w) continue;
            const j = (yy * w + xx) * 4;
            const nr = src[j];
            const ng = src[j + 1];
            const nb = src[j + 2];
            const wgt = weight[Math.abs(nr - r) + Math.abs(ng - g) + Math.abs(nb - b)] * (dx || dy ? 0.8 : 1);
            sr += nr * wgt;
            sg += ng * wgt;
            sb += nb * wgt;
            sw += wgt;
          }
        }
        out[i] = sr / sw + 0.5;
        out[i + 1] = sg / sw + 0.5;
        out[i + 2] = sb / sw + 0.5;
        out[i + 3] = src[i + 3];
      }
    }
    src = out;
  }
  img.data.set(src);
  return img;
}

/**
 * Darken (or, below zero, lighten) toward the corners: nothing in the middle
 * half, rolling off smoothly to the most at the corners.
 */
export function vignette(img, amount) {
  const k = clamp(num(amount) / 100, -1, 1);
  if (!k) return img;
  const { data, width: w, height: h } = img;
  const cx = w / 2;
  const cy = h / 2;
  for (let y = 0; y < h; y++) {
    const dy = (y + 0.5 - cy) / cy;
    for (let x = 0; x < w; x++) {
      const dx = (x + 0.5 - cx) / cx;
      const d = Math.sqrt((dx * dx + dy * dy) / 2);
      const t = clamp((d - 0.35) / 0.65, 0, 1);
      const f = t * t * (3 - 2 * t);
      if (!f) continue;
      const i = (y * w + x) * 4;
      if (k > 0) {
        const m = 1 - k * 0.85 * f;
        data[i] = data[i] * m;
        data[i + 1] = data[i + 1] * m;
        data[i + 2] = data[i + 2] * m;
      } else {
        const m = -k * 0.85 * f;
        data[i] = data[i] + (255 - data[i]) * m;
        data[i + 1] = data[i + 1] + (255 - data[i + 1]) * m;
        data[i + 2] = data[i + 2] + (255 - data[i + 2]) * m;
      }
    }
  }
  return img;
}

/**
 * Red-eye: inside the circle clicked, a pixel much redder than it is green
 * or blue has its red brought down to the others, so a flash-lit pupil goes
 * dark and the rest of the eye, and a red jumper beside it, are left alone.
 * The edge of the circle is feathered so no ring shows.
 */
export function redEye(img, { x, y, radius }) {
  const { data, width: w, height: h } = img;
  const r = Math.max(1, num(radius, 1));
  const x0 = Math.max(0, Math.floor(x - r));
  const x1 = Math.min(w - 1, Math.ceil(x + r));
  const y0 = Math.max(0, Math.floor(y - r));
  const y1 = Math.min(h - 1, Math.ceil(y + r));
  for (let py = y0; py <= y1; py++) {
    for (let px = x0; px <= x1; px++) {
      const d = Math.hypot(px + 0.5 - x, py + 0.5 - y) / r;
      if (d > 1) continue;
      const i = (py * w + px) * 4;
      const red = data[i];
      const other = (data[i + 1] + data[i + 2]) / 2;
      if (red < 50 || red < other * 1.45) continue;
      const strength = d < 0.75 ? 1 : 1 - (d - 0.75) / 0.25;
      // How red decides how much: a pink edge is eased, a red pupil is cleared.
      const redness = clamp((red / Math.max(1, other) - 1.45) / 0.8, 0, 1);
      const k = strength * (0.35 + 0.65 * redness);
      const target = Math.min(other, red) * 0.9;
      data[i] = data[i] + (target - data[i]) * k;
    }
  }
  return img;
}

// ── automatic ───────────────────────────────────────────────────────────────

/** The count of each value per channel, and of luminance. */
export function histogram(img) {
  const r = new Uint32Array(256);
  const g = new Uint32Array(256);
  const b = new Uint32Array(256);
  const l = new Uint32Array(256);
  const { data } = img;
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] === 0) continue;
    r[data[i]]++;
    g[data[i + 1]]++;
    b[data[i + 2]]++;
    l[(0.2126 * data[i] + 0.7152 * data[i + 1] + 0.0722 * data[i + 2] + 0.5) | 0]++;
  }
  return { r, g, b, l };
}

function percentile(counts, fraction) {
  let total = 0;
  for (let i = 0; i < 256; i++) total += counts[i];
  if (!total) return fraction < 0.5 ? 0 : 255;
  const want = total * fraction;
  let seen = 0;
  for (let i = 0; i < 256; i++) {
    seen += counts[i];
    if (seen >= want) return i;
  }
  return 255;
}

/**
 * Auto enhance: the settings a single press should give, worked out from the
 * picture, as Photos' Enhance does. Levels stretch the tones to fill the range
 * (the darkest and brightest half a percent let go), gamma brings the middle
 * tone toward the middle, and a colour cast is half taken out by white
 * balance from the average of the mid-tones (grey world). Each is held back
 * so a picture that is already good is changed little.
 */
export function autoEnhance(img) {
  const hist = histogram(img);
  let inBlack = percentile(hist.l, 0.005);
  let inWhite = percentile(hist.l, 0.995);
  // A picture that already fills the range keeps its points.
  if (inBlack < 6) inBlack = 0;
  if (inWhite > 249) inWhite = 255;
  if (inWhite - inBlack < 32) {
    const mid = (inWhite + inBlack) / 2;
    inBlack = Math.max(0, Math.round(mid - 16));
    inWhite = Math.min(255, Math.round(mid + 16));
  }
  const median = percentile(hist.l, 0.5);
  const m = clamp((median - inBlack) / Math.max(1, inWhite - inBlack), 0.02, 0.98);
  let gamma = Math.log(m) / Math.log(0.5);
  gamma = clamp(1 + (gamma - 1) * 0.6, 0.6, 1.8);

  // Grey world over the mid-tones only: the shadows and highlights carry
  // too little colour, and too much clipping, to say what is grey.
  const { data } = img;
  let sr = 0;
  let sg = 0;
  let sb = 0;
  let n = 0;
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] === 0) continue;
    const l = 0.2126 * data[i] + 0.7152 * data[i + 1] + 0.0722 * data[i + 2];
    if (l < 30 || l > 225) continue;
    sr += data[i];
    sg += data[i + 1];
    sb += data[i + 2];
    n++;
  }
  let temperature = 0;
  let tint = 0;
  if (n) {
    const ar = sr / n;
    const ag = sg / n;
    const ab = sb / n;
    // Half the cast, as a temperature (blue against red) and a tint (green
    // against the two), in the sliders' units.
    temperature = clamp(Math.round(((ab - ar) / Math.max(1, (ar + ab) / 2)) * 0.5 * (100 / 0.56)), -60, 60);
    tint = clamp(Math.round(((ag - (ar + ab) / 2) / Math.max(1, ag)) * 0.5 * (100 / 0.28)), -60, 60);
    if (Math.abs(temperature) < 3) temperature = 0;
    if (Math.abs(tint) < 3) tint = 0;
  }
  return {
    levels: { inBlack, inWhite, gamma: Math.round(gamma * 100) / 100, outBlack: 0, outWhite: 255 },
    temperature,
    tint,
  };
}

/**
 * All of it, in order, on one picture: tone, then noise reduction (before
 * sharpening, or the sharpen would bring the grain back), then sharpen, then
 * the vignette, then red-eye spots (in this copy's pixels).
 */
export function applyPixels(img, settings = {}, { spots = [] } = {}) {
  const s = { ...DEFAULT_PIXELS, ...settings };
  const needsTone = s.exposure || s.highlights || s.shadows || s.temperature || s.tint || (s.levels && !isNeutral({ levels: s.levels })) || (s.curve && !isNeutral({ curve: s.curve }));
  if (needsTone) applyTone(img, s);
  if (s.denoise) denoise(img, s.denoise);
  // One pixel of radius for every 1,600 of the longer side of the copy
  // being worked, so a preview at a third of the size gets a third.
  if (s.sharpen) sharpen(img, s.sharpen, { radius: Math.max(1, Math.round(Math.max(img.width, img.height) / 1600)) });
  if (s.vignette) vignette(img, s.vignette);
  for (const spot of spots) redEye(img, spot);
  return img;
}
