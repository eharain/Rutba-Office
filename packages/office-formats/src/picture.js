// SPDX-License-Identifier: AGPL-3.0-or-later
//
// A picture kept inside a document, made something a window can draw.
//
// Word, Excel and PowerPoint keep a picture as the file it came from. Most
// are PNG or JPEG, which a window draws as they are. Some are not: a chart
// or a diagram pasted from another program is kept as an EMF or a WMF, a
// list of drawing calls no browser reads. Documents, Worksheets and
// Presentations all ask this one place, so a picture that draws in one app
// draws in the other two.
//
// drawablePicture(bytes, name) answers { type, bytes }:
//   - an EMF or a WMF, played as an SVG (emf.js, wmf.js); an EMF that is
//     only bitmaps the SVG could not be made of, as a PNG (emfToPng);
//   - a picture a window draws as it is, unchanged;
//   - any other kind, through the decoder lent with setPictureDecoder;
//   - null when none of these can draw it, so the caller shows the
//     picture's empty frame rather than a broken one.
//
// The answer is kept against the bytes it was made from, so a picture drawn
// once per page or per slide is played once. Nothing here throws: a damaged
// picture is null.

import zlib from 'node:zlib';
import { emfToSvg, emfToPng, isEmf } from './emf.js';
import { wmfToSvg, isWmf } from './wmf.js';

/** What a window draws as it is, by extension. */
const WINDOW_DRAWS = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', dib: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml', ico: 'image/x-icon',
};

const extOf = (name) => String(name ?? '').toLowerCase().split(/[\\/]/).pop().split('.').pop();

let lentDecoder = null;

/**
 * Lend a decoder for picture kinds a window does not draw (TIFF, HEIC and
 * the rest): (bytes, name) => ({ type, bytes } | null). Pass null to stop.
 * It is asked after the metafiles, and may throw: a throw is a null.
 */
export function setPictureDecoder(fn) {
  lentDecoder = typeof fn === 'function' ? fn : null;
  cache = new WeakMap();
}

let cache = new WeakMap();

/** Is this picture a metafile, by its bytes first and then its name? */
export function metafileKind(bytes, name) {
  if (bytes && bytes.length >= 4) {
    try {
      const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer ?? bytes, bytes.byteOffset ?? 0, bytes.byteLength ?? bytes.length);
      if (isEmf(b)) return 'emf';
      if (isWmf(b)) return 'wmf';
    } catch { /* not one */ }
  }
  const ext = extOf(name);
  return ext === 'emf' || ext === 'emz' ? 'emf' : ext === 'wmf' || ext === 'wmz' ? 'wmf' : null;
}

function play(kind, b) {
  if (kind === 'emf') {
    const drawn = emfToSvg(b);
    if (drawn) return { type: 'image/svg+xml', bytes: Buffer.from(drawn.svg, 'utf8'), width: drawn.width, height: drawn.height };
    const flat = emfToPng(b);
    if (flat) return { type: 'image/png', bytes: Buffer.from(flat.png), width: flat.width, height: flat.height };
    return null;
  }
  const drawn = wmfToSvg(b);
  return drawn ? { type: 'image/svg+xml', bytes: Buffer.from(drawn.svg, 'utf8'), width: drawn.width, height: drawn.height } : null;
}

/**
 * A picture's bytes as a window can draw them.
 * @param {Uint8Array} bytes the picture part, as the document keeps it
 * @param {string} [name] its part name or file name, for its extension
 * @returns {{ type: string, bytes: Uint8Array, width?: number, height?: number } | null}
 */
export function drawablePicture(bytes, name = '') {
  if (!bytes || !bytes.length) return null;
  const keyed = typeof bytes === 'object';
  if (keyed && cache.has(bytes)) return cache.get(bytes);
  let out = null;
  try {
    let b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer ?? bytes, bytes.byteOffset ?? 0, bytes.byteLength ?? bytes.length);
    // An .emz or a .wmz is the metafile gzipped.
    if (b[0] === 0x1f && b[1] === 0x8b && /^(emz|wmz)$/.test(extOf(name))) b = zlib.gunzipSync(b, { maxOutputLength: 64 * 1048576 });
    const kind = metafileKind(b, name);
    if (kind) out = play(kind, b);
    else if (WINDOW_DRAWS[extOf(name)]) out = { type: WINDOW_DRAWS[extOf(name)], bytes: b };
    else if (lentDecoder) {
      const lent = lentDecoder(b, name);
      out = lent && lent.bytes && lent.type ? lent : null;
    }
  } catch {
    out = null;
  }
  if (keyed) cache.set(bytes, out);
  return out;
}

/** The same, as a data: URI an SVG or a page can carry, or null. */
export function pictureDataUri(bytes, name = '') {
  const p = drawablePicture(bytes, name);
  return p ? `data:${p.type};base64,${Buffer.from(p.bytes).toString('base64')}` : null;
}

/** Is a picture of this name one this module changes before a window draws it? */
export function needsDrawing(name) {
  const ext = extOf(name);
  return ext === 'emf' || ext === 'wmf' || ext === 'emz' || ext === 'wmz' || (!WINDOW_DRAWS[ext] && lentDecoder != null);
}

export default { drawablePicture, pictureDataUri, setPictureDecoder, metafileKind, needsDrawing };
