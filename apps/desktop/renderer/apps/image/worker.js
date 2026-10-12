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
import { FORMATS, sizeFor, withExifFrom, underSize } from '@rutba/imaging/batch';
import { readExif } from '@rutba/imaging/exif';
import { probeImage } from '@rutba/imaging/probe';

async function decode(bytes, upright) {
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
  const { format = 'jpeg', quality = 0.9, resize = { mode: 'none' }, upright = true, strip = true, maxBytes = 0 } = options;
  const info = FORMATS[format];
  if (!info) throw new Error(`unknown format ${format}`);
  const bitmap = await decode(bytes, upright);
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
  const encode = async (q) => new Uint8Array(await (await canvas.convertToBlob({ type: info.mime, quality: info.quality ? q : undefined })).arrayBuffer());
  const result = info.quality && maxBytes
    ? await underSize(encode, maxBytes, { high: quality })
    : { bytes: await encode(quality), quality, fits: true };
  let out = result.bytes;
  if (!strip && format === 'jpeg') {
    // The camera's details kept, the orientation upright if the pixels are.
    const stored = probeImage(bytes);
    const turned = stored && stored.width && (stored.width !== size.width || stored.height !== size.height) && (stored.width / stored.height > 1) !== (size.width / size.height > 1);
    const orientation = readExif(bytes).orientation || 1;
    out = withExifFrom(out, bytes, { orientation: upright || turned ? 1 : orientation });
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
    } else {
      self.postMessage({ id: job.id, error: `unknown job ${job.kind}` });
    }
  } catch (err) {
    self.postMessage({ id: job.id, error: err?.message || String(err) });
  }
};
