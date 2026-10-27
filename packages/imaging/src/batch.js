// Batch work over many pictures, and the parts of saving one picture that are
// arithmetic rather than pixels: the size a picture ends up, the name a
// pattern gives it, a free name beside a file that is already there, the
// metadata a JPEG keeps, and the quality that brings a file under a size.
//
// No DOM and no file system: the window (or the image worker) decodes and
// encodes, and asks here what to call the result and how large to make it.

import { withOrientation } from './exif.js';

/** The formats a picture is written in, by the name the windows give them. */
export const FORMATS = {
  png: { mime: 'image/png', ext: 'png', quality: false },
  jpeg: { mime: 'image/jpeg', ext: 'jpg', quality: true },
  webp: { mime: 'image/webp', ext: 'webp', quality: true },
  // Written by the suite's own writers (encode.js), not a canvas.
  bmp: { mime: 'image/bmp', ext: 'bmp', quality: false, own: true },
  tiff: { mime: 'image/tiff', ext: 'tif', quality: false, own: true },
  ico: { mime: 'image/x-icon', ext: 'ico', quality: false, own: true },
};

/** The format a file's own extension says, or null for one we do not write. */
export function formatOf(fileName) {
  const ext = String(fileName || '').toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] || '';
  if (ext === 'png') return 'png';
  if (ext === 'jpg' || ext === 'jpeg' || ext === 'jpe') return 'jpeg';
  if (ext === 'webp') return 'webp';
  if (ext === 'bmp') return 'bmp';
  if (ext === 'tif' || ext === 'tiff') return 'tiff';
  if (ext === 'ico') return 'ico';
  return null;
}

/**
 * The size a picture is written at.
 *   { mode: 'none' }                       as it is
 *   { mode: 'percent', percent }           a share of it
 *   { mode: 'fit', width, height }         inside the box, the ratio kept, never made larger
 *   { mode: 'long', size }                 the longer side at most this, the ratio kept
 *   { mode: 'exact', width, height }       exactly so
 */
export function sizeFor(size, resize = {}) {
  const w = size.width;
  const h = size.height;
  const round = (v) => Math.max(1, Math.round(v));
  switch (resize.mode) {
    case 'percent': {
      const p = Math.max(1, Math.min(1000, Number(resize.percent) || 100)) / 100;
      return { width: round(w * p), height: round(h * p) };
    }
    case 'fit': {
      const bw = Number(resize.width) || w;
      const bh = Number(resize.height) || h;
      const s = Math.min(1, bw / w, bh / h);
      return { width: round(w * s), height: round(h * s) };
    }
    case 'long': {
      const cap = Number(resize.size) || Math.max(w, h);
      const s = Math.min(1, cap / Math.max(w, h));
      return { width: round(w * s), height: round(h * s) };
    }
    case 'exact':
      return { width: round(Number(resize.width) || w), height: round(Number(resize.height) || h) };
    default:
      return { width: w, height: h };
  }
}

const BAD = /[\\/:*?"<>|\u0000-\u001f]/g;

/**
 * A file name from a pattern, without its extension.
 *   {name}  the original's name without its extension
 *   {n}     the picture's place in the batch, from 1; {n:3} pads to three digits
 *   {date}  the file's date, 2026-10-12
 *   {w} {h} the size it is written at
 * Characters no file system takes are dropped; an empty result is the
 * original's name.
 */
export function nameFromPattern(pattern, { name = 'picture', n = 1, date = null, width = 0, height = 0 } = {}) {
  const day = date ? new Date(date) : null;
  const iso = day && !Number.isNaN(day.getTime())
    ? `${day.getFullYear()}-${String(day.getMonth() + 1).padStart(2, '0')}-${String(day.getDate()).padStart(2, '0')}`
    : '';
  const out = String(pattern || '{name}')
    // Each through a function, so a $ in a file's name is taken as it is.
    .replace(/\{name\}/g, () => name)
    .replace(/\{n(?::(\d))?\}/g, (_m, pad) => String(n).padStart(Number(pad) || 1, '0'))
    .replace(/\{date\}/g, () => iso)
    .replace(/\{w\}/g, () => String(width))
    .replace(/\{h\}/g, () => String(height))
    .replace(BAD, '')
    .replace(/[. ]+$/, '')
    .trim();
  return out || name;
}

const splitName = (file) => {
  const base = String(file).split(/[\\/]/).pop();
  const dot = base.lastIndexOf('.');
  return dot > 0 ? { stem: base.slice(0, dot), ext: base.slice(dot + 1) } : { stem: base, ext: '' };
};

const join = (dir, file) => (dir ? `${String(dir).replace(/[\\/]+$/, '')}${String(dir).includes('\\') ? '\\' : '/'}${file}` : file);

/** The folder a path is in, as the path writes it. */
export function folderOf(file) {
  const at = Math.max(String(file).lastIndexOf('/'), String(file).lastIndexOf('\\'));
  return at < 0 ? '' : String(file).slice(0, at);
}

/**
 * A name beside `taken` (a function saying whether a path is used) that is
 * free: "photo.jpg", then "photo (2).jpg", "photo (3).jpg", as Explorer
 * names a copy.
 */
export function freeName(file, taken) {
  if (!taken(file)) return file;
  const dir = folderOf(file);
  const { stem, ext } = splitName(file);
  for (let i = 2; i < 10000; i++) {
    const next = join(dir, `${stem} (${i})${ext ? `.${ext}` : ''}`);
    if (!taken(next)) return next;
  }
  return file;
}

/**
 * What a batch writes: for each picture, where and in what format, and the
 * pictures whose result would land on a file already there (`clashes`), so
 * the window can ask before anything is replaced. Two pictures of the batch
 * that would get the same name are told apart with " (2)" as they go; a
 * result never lands on its own original unless `replace` says so.
 *
 * files:   [{ path, name?, size: { width, height }, date? }]
 * options: { format: 'same' | 'png' | 'jpeg' | 'webp', pattern, folder (null: beside each),
 *            resize, exists (a function: is this path a file now) }
 */
export function planBatch(files, { format = 'same', pattern = '{name}', folder = null, resize = { mode: 'none' }, exists = () => false } = {}) {
  const planned = new Set();
  const items = [];
  files.forEach((file, i) => {
    const { stem } = splitName(file.path);
    const fmt = format === 'same' ? formatOf(file.path) || 'png' : format;
    const size = sizeFor(file.size || { width: 0, height: 0 }, resize);
    const name = nameFromPattern(pattern, { name: stem, n: i + 1, date: file.date, width: size.width, height: size.height });
    const dir = folder || folderOf(file.path);
    let target = join(dir, `${name}.${FORMATS[fmt].ext}`);
    // Within the batch, a later picture never takes an earlier one's name.
    target = freeName(target, (p) => planned.has(p.toLowerCase()));
    planned.add(target.toLowerCase());
    const original = target.toLowerCase() === String(file.path).toLowerCase();
    items.push({ source: file.path, target, format: fmt, size, clash: original || Boolean(exists(target)), original });
  });
  return { items, clashes: items.filter((it) => it.clash) };
}

/**
 * The plan with every clash moved to a free name ("Keep both"): the
 * originals and whatever else is in the folder are left as they are.
 */
export function keepBoth(plan, exists) {
  const used = new Set(plan.items.map((it) => it.target.toLowerCase()));
  const items = plan.items.map((it) => {
    if (!it.clash) return it;
    const target = freeName(it.target, (p) => used.has(p.toLowerCase()) || exists(p));
    used.add(target.toLowerCase());
    return { ...it, target, clash: false, original: false };
  });
  return { items, clashes: [] };
}

// ── JPEG metadata ───────────────────────────────────────────────────────────

/** The APP1 Exif segment of a JPEG, whole (marker and length included), or null. */
export function exifSegment(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  let at = 2;
  while (at + 4 <= b.length && b[at] === 0xff) {
    const marker = b[at + 1];
    const length = (b[at + 2] << 8) | b[at + 3];
    if (length < 2) return null;
    if (marker === 0xe1 && at + 10 <= b.length && String.fromCharCode(b[at + 4], b[at + 5], b[at + 6], b[at + 7]) === 'Exif') {
      return b.slice(at, Math.min(b.length, at + 2 + length));
    }
    if (marker === 0xda) break;
    at += 2 + length;
  }
  return null;
}

/**
 * A JPEG with nothing but its picture: every APPn segment but the JFIF
 * header, and every comment, left out, and the camera, date, place and thumbnail
 * with them.
 */
export function stripJpeg(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return b;
  const parts = [b.subarray(0, 2)];
  let at = 2;
  while (at + 4 <= b.length && b[at] === 0xff) {
    const marker = b[at + 1];
    if (marker === 0xda) break;
    const length = (b[at + 2] << 8) | b[at + 3];
    if (length < 2) return b;
    const keep = !((marker >= 0xe1 && marker <= 0xef) || marker === 0xfe);
    if (keep) parts.push(b.subarray(at, at + 2 + length));
    at += 2 + length;
  }
  parts.push(b.subarray(at));
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

/**
 * A freshly encoded JPEG given the original's Exif block, its orientation
 * set to 1 because the pixels written are already upright (or to the one
 * given, where they were written as stored). Answers the JPEG unchanged
 * where the original had no Exif.
 */
export function withExifFrom(encoded, original, { orientation = 1 } = {}) {
  const segment = original ? exifSegment(original) : null;
  const b = encoded instanceof Uint8Array ? encoded : new Uint8Array(encoded);
  if (!segment || b[0] !== 0xff || b[1] !== 0xd8) return b;
  // After the JFIF header, where a camera puts it.
  let after = 2;
  if (b[2] === 0xff && b[3] === 0xe0) after = 4 + ((b[4] << 8) | b[5]);
  const out = new Uint8Array(b.length + segment.length);
  out.set(b.subarray(0, after), 0);
  out.set(segment, after);
  out.set(b.subarray(after), after + segment.length);
  return withOrientation(out, orientation) || out;
}

// ── a size cap ──────────────────────────────────────────────────────────────

/**
 * The best quality whose file fits under `maxBytes`, by halving the range:
 * `encode(quality)` answers the bytes at that quality (0.05 to 1). Answers
 * `{ bytes, quality, fits }`; when even the lowest does not fit, the lowest,
 * with `fits: false`, so the window can say so rather than pretend.
 */
export async function underSize(encode, maxBytes, { high = 0.92, low = 0.05, steps = 7 } = {}) {
  let best = await encode(high);
  if (!maxBytes || best.length <= maxBytes) return { bytes: best, quality: high, fits: true };
  let lo = low;
  let hi = high;
  let found = null;
  for (let i = 0; i < steps; i++) {
    const q = Math.round(((lo + hi) / 2) * 100) / 100;
    const bytes = await encode(q);
    if (bytes.length <= maxBytes) {
      found = { bytes, quality: q, fits: true };
      lo = q;
    } else {
      hi = q;
      best = bytes;
    }
  }
  if (found) return found;
  const bytes = await encode(low);
  return { bytes, quality: low, fits: bytes.length <= maxBytes };
}
