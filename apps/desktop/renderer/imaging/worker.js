// The image worker: the pixel work that would make a window stutter, done
// on a thread of its own. Bundled on its own as image-worker.js and served
// from the app like the window's own code; it has no bridge to the shell and
// reads no file: it is handed pixels or a file's bytes and hands back the
// same.
//
//   { kind: 'pixels', width, height, buffer, settings, spots }
//       → { buffer }                        the picture with tone and detail applied
//   { kind: 'convert', bytes, name, options }
//       → { bytes, width, height, format, quality, fits }   one picture of a batch
import { applyPixels } from '@rutba/imaging/pixels';
import { FORMATS, sizeFor, underSize } from '@rutba/imaging/batch';
import { readExif, keepExif } from '@rutba/imaging/exif';
import { probeImage } from '@rutba/imaging/probe';
import { decodePicture } from '@rutba/imaging/decode';
import { encodePicture, encodeIco, iconSizes } from '@rutba/imaging/encode';
import { runTool } from './tool-jobs.js';

async function decode(bytes, upright, name = '') {
  try {
    return await decodeHere(bytes, upright);
  } catch (err) {
    // A format Chromium does not draw (a TIFF, say): read by the suite's own readers.
    let pic;
    try {
      pic = decodePicture(bytes, name);
    } catch (refusal) {
      throw refusal?.picture ? refusal : err;
    }
    return createImageBitmap(new ImageData(new Uint8ClampedArray(pic.rgba.buffer, pic.rgba.byteOffset, pic.rgba.byteLength), pic.width, pic.height));
  }
}

/** BMP, TIFF and ICO, which a canvas does not write, by the suite's own writers. */
async function ownBytes(canvas, format) {
  if (format === 'ico') {
    const entries = [];
    for (const size of iconSizes(canvas.width, canvas.height)) {
      const c = new OffscreenCanvas(size, size);
      const g = c.getContext('2d');
      g.imageSmoothingEnabled = true;
      g.imageSmoothingQuality = 'high';
      const k = Math.min(size / canvas.width, size / canvas.height);
      const w = Math.max(1, Math.round(canvas.width * k));
      const h = Math.max(1, Math.round(canvas.height * k));
      g.drawImage(canvas, Math.round((size - w) / 2), Math.round((size - h) / 2), w, h);
      entries.push({ width: size, height: size, png: new Uint8Array(await (await c.convertToBlob({ type: 'image/png' })).arrayBuffer()) });
    }
    return encodeIco(entries);
  }
  const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
  return encodePicture(new Uint8Array(data.data.buffer, data.data.byteOffset, data.data.byteLength), canvas.width, canvas.height, format);
}

async function decodeHere(bytes, upright) {
  const blob = new Blob([bytes]);
  if (!upright) {
    // As stored, with no turn from its EXIF; where the platform cannot say
    // that, as it shows it, which the caller can tell by the size.
    try {
      return await createImageBitmap(blob, { imageOrientation: 'none' });
    } catch {
      /* fall through */
    }
  }
  return createImageBitmap(blob, { imageOrientation: 'from-image' });
}

async function convert(bytes, options = {}) {
  const { format = 'jpeg', quality = 0.9, resize = { mode: 'none' }, upright = true, strip = true, maxBytes = 0, name = '' } = options;
  const info = FORMATS[format];
  if (!info) throw new Error(`unknown format ${format}`);
  const bitmap = await decode(bytes, upright, name);
  const size = sizeFor({ width: bitmap.width, height: bitmap.height }, resize);
  const canvas = new OffscreenCanvas(size.width, size.height);
  const ctx = canvas.getContext('2d');
  if (format === 'jpeg') {
    // A JPEG has no transparency: what was clear becomes white, not black.
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, size.width, size.height);
  }
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(bitmap, 0, 0, size.width, size.height);
  bitmap.close?.();
  const encode = async (q) => (info.own ? ownBytes(canvas, format) : new Uint8Array(await (await canvas.convertToBlob({ type: info.mime, quality: info.quality ? q : undefined })).arrayBuffer()));
  const result = info.quality && maxBytes
    ? await underSize(encode, maxBytes, { high: quality })
    : { bytes: await encode(quality), quality, fits: true };
  let out = result.bytes;
  // strip: true (or 'all') drops every detail; 'location' keeps the camera
  // and drops the place; false keeps both.
  if (strip !== true && strip !== 'all') {
    // The camera's details kept, the orientation upright if the pixels are,
    // in a JPEG's APP1, a PNG's eXIf or a WebP's EXIF chunk.
    const stored = probeImage(bytes);
    const turned = stored && stored.width && (stored.width !== size.width || stored.height !== size.height) && (stored.width / stored.height > 1) !== (size.width / size.height > 1);
    const orientation = readExif(bytes).orientation || 1;
    out = keepExif(out, format, bytes, { orientation: upright || turned ? 1 : orientation, width: size.width, height: size.height, location: strip !== 'location' });
  }
  if (maxBytes && !info.quality) result.fits = out.length <= maxBytes;
  return { bytes: out, width: size.width, height: size.height, format, quality: result.quality, fits: result.fits };
}

self.onmessage = async (e) => {
  const job = e.data || {};
  try {
    if (job.kind === 'pixels') {
      const data = new Uint8ClampedArray(job.buffer);
      const img = { data, width: job.width, height: job.height };
      applyPixels(img, job.settings || {}, { spots: job.spots || [] });
      self.postMessage({ id: job.id, buffer: data.buffer }, [data.buffer]);
    } else if (job.kind === 'convert') {
      const done = await convert(job.bytes, job.options);
      self.postMessage({ id: job.id, ...done }, [done.bytes.buffer]);
    } else if (job.kind === 'tool') {
      const done = runTool(job);
      self.postMessage({ id: job.id, ...done.result }, done.transfer);
    } else {
      self.postMessage({ id: job.id, error: `unknown job ${job.kind}` });
    }
  } catch (err) {
    self.postMessage({ id: job.id, error: err?.message || String(err) });
  }
};
