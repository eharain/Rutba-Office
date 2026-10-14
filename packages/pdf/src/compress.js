'use strict';

/**
 * A smaller PDF: its large pictures drawn again at the resolution they are
 * shown at, as Acrobat's Reduce File Size does.
 *
 * Each picture's size on its pages is found by reading every page's content
 * with its transformation matrix (and the forms it draws), so a picture of
 * 3000 pixels shown 3 inches wide is seen to be at 1000 dots an inch. One
 * whose resolution is over the target, and whose bytes are worth the work,
 * is planned for: its samples (decoded here when this module can, else the
 * file's JPEG bytes for the window to decode), its size now and the size it
 * needs. The caller hands back each picture encoded smaller (JPEG from the
 * window's canvas), or this module resamples a picture it decoded and
 * deflates it. The file is then written whole by assemble.js with the new
 * pictures in place of the old, which are no longer in it.
 *
 * Left alone: masks and pictures with a /Decode array, pictures of more
 * than three colour channels (CMYK, which a canvas does not keep), and any
 * picture a replacement would not make smaller.
 */

const zlib = require('zlib');
const { openPdf, Name, Ref, Dict, Stream } = require('./reader');
const { assemble } = require('./assemble');
const { parseContent } = require('./redact');

const num = (v, fallback = 0) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);
const mul = (a, b) => [
  a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5],
];

/** The largest size, in points, each picture object is drawn at: Map(num -> { w, h }). */
function drawnSizes(file) {
  const sizes = new Map();
  const seenForms = new Set();
  const walk = (data, resources, base, depth) => {
    if (depth > 12) return;
    const res = file.resolve(resources);
    const xo = res instanceof Dict ? file.resolve(res.get('XObject')) : null;
    let ctm = base.slice();
    const stack = [];
    let ops;
    try { ops = parseContent(data); } catch { return; }
    for (const o of ops) {
      if (o.op === 'q') stack.push(ctm.slice());
      else if (o.op === 'Q') { if (stack.length) ctm = stack.pop(); }
      else if (o.op === 'cm' && o.operands.length >= 6) ctm = mul(o.operands.slice(0, 6).map((x) => num(x)), ctm);
      else if (o.op === 'Do' && o.operands[0] instanceof Name && xo instanceof Dict) {
        const ref = xo.get(o.operands[0].name);
        const x = file.resolve(ref);
        if (!(x instanceof Stream) || !(ref instanceof Ref)) continue;
        const sub = x.dict.get('Subtype') instanceof Name ? x.dict.get('Subtype').name : '';
        if (sub === 'Image') {
          const w = Math.hypot(ctm[0], ctm[1]);
          const h = Math.hypot(ctm[2], ctm[3]);
          const had = sizes.get(ref.num) || { w: 0, h: 0 };
          sizes.set(ref.num, { w: Math.max(had.w, w), h: Math.max(had.h, h) });
        } else if (sub === 'Form') {
          const key = `${ref.num}@${ctm.join(',')}`;
          if (seenForms.has(key)) continue;
          seenForms.add(key);
          const fm = file.resolve(x.dict.get('Matrix'));
          const m = Array.isArray(fm) && fm.length >= 6 ? fm.slice(0, 6).map((v) => num(file.resolve(v))) : [1, 0, 0, 1, 0, 0];
          let inner;
          try { inner = file.decode(x); } catch { continue; }
          walk(inner, x.dict.get('Resources') || resources, mul(m, ctm), depth + 1);
        }
      }
    }
  };
  file.pages().forEach((page, p) => {
    const resources = page.dict.has('Resources') ? page.dict.get('Resources') : page.inherited.Resources;
    walk(file.pageContent(p), resources, [1, 0, 0, 1, 0, 0], 0);
  });
  return sizes;
}

/** An image's samples when this module can read them (no compression or Flate, 8 bits, grey or RGB). */
function samples(file, stream) {
  const d = stream.dict;
  const filters = file.resolve(d.get('Filter'));
  const list = (Array.isArray(filters) ? filters : filters ? [filters] : []).map((f) => (f instanceof Name ? f.name : ''));
  if (list.some((f) => !['FlateDecode', 'Fl'].includes(f))) return null;
  if (num(file.resolve(d.get('BitsPerComponent')), 8) !== 8) return null;
  const cs = file.resolve(d.get('ColorSpace'));
  const comps = cs instanceof Name ? ({ DeviceRGB: 3, RGB: 3, DeviceGray: 1, G: 1 }[cs.name] || 0) : 0;
  if (!comps) return null;
  const width = num(file.resolve(d.get('Width')));
  const height = num(file.resolve(d.get('Height')));
  let data;
  try { data = file.decode(stream); } catch { return null; }
  if (data.length < width * height * comps) return null;
  return { width, height, comps, data: data.subarray(0, width * height * comps) };
}

/**
 * The pictures worth making smaller for a target resolution: [{ ref, width,
 * height, toWidth, toHeight, filter, comps, data, bytes }], `data` the
 * JPEG's bytes for a JPEG and the decoded samples for one decoded here.
 */
function compressPlan(bytes, { dpi = 150, minBytes = 20000 } = {}) {
  const file = openPdf(bytes);
  const sizes = drawnSizes(file);
  const out = [];
  for (const [n, size] of sizes) {
    const x = file.getObject(n);
    if (!(x instanceof Stream)) continue;
    const d = x.dict;
    if (file.resolve(d.get('ImageMask')) === true || d.get('Decode')) continue;
    const width = num(file.resolve(d.get('Width')));
    const height = num(file.resolve(d.get('Height')));
    if (!(width > 0 && height > 0) || x.data.length < minBytes) continue;
    // Pixels needed to show it at the target resolution where it is drawn largest.
    const toWidth = Math.max(1, Math.round((size.w / 72) * dpi));
    const toHeight = Math.max(1, Math.round((size.h / 72) * dpi));
    if (toWidth >= width * 0.9 && toHeight >= height * 0.9) continue;
    const f = file.resolve(d.get('Filter'));
    const filter = f instanceof Name ? f.name : Array.isArray(f) ? f.map((q) => (q instanceof Name ? q.name : '')).join(' ') : '';
    if (/DCTDecode|DCT/.test(filter)) {
      const cs = file.resolve(d.get('ColorSpace'));
      const comps = cs instanceof Name ? ({ DeviceRGB: 3, DeviceGray: 1 }[cs.name] || 0) : Array.isArray(cs) && cs[0] instanceof Name && cs[0].name === 'ICCBased' ? num(file.resolve(file.resolve(cs[1])?.dict?.get('N')), 0) : 0;
      if (comps !== 1 && comps !== 3) continue;
      out.push({ ref: n, width, height, toWidth: Math.min(width, toWidth), toHeight: Math.min(height, toHeight), filter, comps, data: x.data, bytes: x.data.length });
      continue;
    }
    const pix = samples(file, x);
    if (!pix) continue;
    out.push({ ref: n, width, height, toWidth: Math.min(width, toWidth), toHeight: Math.min(height, toHeight), filter, comps: pix.comps, data: Buffer.from(pix.data), bytes: x.data.length });
  }
  return out;
}

/** Samples made smaller by averaging each output pixel's box of input pixels. */
function resample({ width, height, comps, data }, toWidth, toHeight) {
  const out = Buffer.alloc(toWidth * toHeight * comps);
  const sx = width / toWidth;
  const sy = height / toHeight;
  for (let y = 0; y < toHeight; y++) {
    const y0 = Math.floor(y * sy);
    const y1 = Math.max(y0 + 1, Math.floor((y + 1) * sy));
    for (let x = 0; x < toWidth; x++) {
      const x0 = Math.floor(x * sx);
      const x1 = Math.max(x0 + 1, Math.floor((x + 1) * sx));
      for (let c = 0; c < comps; c++) {
        let sum = 0;
        for (let yy = y0; yy < y1; yy++) for (let xx = x0; xx < x1; xx++) sum += data[(yy * width + xx) * comps + c];
        out[(y * toWidth + x) * comps + c] = Math.round(sum / ((y1 - y0) * (x1 - x0)));
      }
    }
  }
  return out;
}

/**
 * The file written whole with its large pictures made smaller. `images`:
 * { [ref]: { data, width, height, filter: 'DCTDecode', comps } } handed back
 * encoded by the caller; a planned picture decoded here and not handed back
 * is resampled and deflated here.
 * @returns {{ bytes: Buffer, replaced: number, before: number, after: number }}
 */
function compressPdf(bytes, { dpi = 150, images = {}, title } = {}) {
  const file = openPdf(bytes);
  if (file.encrypted) throw new Error('A PDF protected with a password cannot be made smaller yet.');
  const plan = compressPlan(bytes, { dpi });
  const replace = new Map();
  for (const item of plan) {
    const old = file.getObject(item.ref);
    const given = images[item.ref];
    let stream = null;
    if (given && given.data && given.data.length) {
      const dict = new Dict();
      for (const [k, v] of old.dict.entries()) if (!['Filter', 'DecodeParms', 'Length', 'DL', 'Width', 'Height', 'ColorSpace', 'BitsPerComponent'].includes(k)) dict.set(k, v);
      dict.set('Width', given.width);
      dict.set('Height', given.height);
      dict.set('ColorSpace', new Name((given.comps || 3) === 1 ? 'DeviceGray' : 'DeviceRGB'));
      dict.set('BitsPerComponent', 8);
      dict.set('Filter', new Name(given.filter || 'DCTDecode'));
      stream = new Stream(dict, Buffer.from(given.data));
    } else if (!/DCT/.test(item.filter)) {
      const data = resample({ width: item.width, height: item.height, comps: item.comps, data: item.data }, item.toWidth, item.toHeight);
      const dict = new Dict();
      for (const [k, v] of old.dict.entries()) if (!['Filter', 'DecodeParms', 'Length', 'DL', 'Width', 'Height'].includes(k)) dict.set(k, v);
      dict.set('Width', item.toWidth);
      dict.set('Height', item.toHeight);
      dict.set('Filter', new Name('FlateDecode'));
      stream = new Stream(dict, zlib.deflateSync(data, { level: 9 }));
    }
    if (stream && stream.data.length < old.data.length) replace.set(item.ref, stream);
  }
  const items = file.pages().map((_, page) => ({ file, page }));
  const out = assemble(items, { title, replace: new Map([[file, replace]]) });
  return { bytes: out, replaced: replace.size, before: bytes.length, after: out.length };
}

module.exports = { compressPlan, compressPdf, drawnSizes, resample };
