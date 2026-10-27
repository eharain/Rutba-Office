// The pictures Image opens and writes beyond what Chromium does itself.
//
// Opening: a format Chromium cannot draw (TIFF, and the rest decodePicture
// reads) arrives over rutba://file already decoded, a page at a time (the
// main process's picture thread does it), and Image takes the PNG of the
// page it was asked for. Writing: BMP, TIFF and ICO are made here from the
// canvas's pixels by the suite's own writers (@rutba/imaging/encode), since
// a canvas writes only PNG, JPEG and WebP.

import { needsDecoding } from '@rutba/imaging/decode';
import { encodePicture, encodeIco, iconSizes } from '@rutba/imaging/encode';
import { t } from '@rutba/office-ui';

/** The extensions Image's Open offers, its project first. */
export const IMAGE_OPENS = ['rimg', 'png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'avif', 'tif', 'tiff', 'ico', 'cur', 'svg', 'qoi', 'pbm', 'pgm', 'ppm', 'pnm', 'pam', 'tga', 'pcx', 'hdr', 'psd', 'psb'];

/** A file Image takes decoded from the file host rather than from its bytes. */
export const servedDecoded = (name) => needsDecoding(name);

/** Files with pages. */
export const PAGED = /\.(tiff?)$/i;

/** The address of one page of a file. */
export function pageUrl(url, page = 0) {
  return page ? `${url}${url.includes('?') ? '&' : '?'}page=${page}` : url;
}

/** What the file host says of a file it decodes: `{ pages, width, height }` or `{ error }`. */
export async function servedInfo(url) {
  try {
    return await (await fetch(`${url}${url.includes('?') ? '&' : '?'}info=1`)).json();
  } catch (err) {
    return { error: err?.message || String(err) };
  }
}

/** The PNG of a page of a file the host decodes; a refusal comes back as its sentence. */
export async function servedPng(url, page = 0) {
  const res = await fetch(pageUrl(url, page));
  if (!res.ok) throw new Error((await res.text().catch(() => '')) || t('This file could not be decoded as an image.'));
  return new Uint8Array(await res.arrayBuffer());
}

const blobBytes = async (blob) => new Uint8Array(await blob.arrayBuffer());

/**
 * A canvas written as `format`: a canvas's own encoder for PNG, JPEG and
 * WebP, the suite's for BMP, TIFF and ICO. An icon is drawn at each size it
 * carries, fitted and centred on a clear square, with the canvas's own
 * smoothing, and each size kept as a PNG inside it.
 */
export async function canvasBytes(canvas, info, format, quality) {
  if (!info.own) {
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, info.mime, info.quality ? quality : undefined));
    if (!blob) throw new Error(t('The picture could not be written.'));
    return blobBytes(blob);
  }
  if (format === 'ico') {
    const entries = [];
    for (const size of iconSizes(canvas.width, canvas.height)) {
      const c = document.createElement('canvas');
      c.width = size;
      c.height = size;
      const g = c.getContext('2d');
      g.imageSmoothingEnabled = true;
      g.imageSmoothingQuality = 'high';
      const k = Math.min(size / canvas.width, size / canvas.height);
      const w = Math.max(1, Math.round(canvas.width * k));
      const h = Math.max(1, Math.round(canvas.height * k));
      g.drawImage(canvas, Math.round((size - w) / 2), Math.round((size - h) / 2), w, h);
      const blob = await new Promise((resolve) => c.toBlob(resolve, 'image/png'));
      entries.push({ width: size, height: size, png: await blobBytes(blob) });
    }
    return encodeIco(entries);
  }
  const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
  const rgba = new Uint8Array(data.data.buffer, data.data.byteOffset, data.data.byteLength);
  return encodePicture(rgba, canvas.width, canvas.height, format);
}

/** What each format the Export tab offers keeps, said under the quality slider. */
export function formatHint(format, info) {
  if (info?.quality) return t('Lower is a smaller file.');
  if (format === 'bmp') return t('BMP keeps every pixel, unpacked.');
  if (format === 'tiff') return t('TIFF keeps every pixel, for print and scans.');
  if (format === 'ico') return t('An icon, at each size from 256 down to 16 pixels.');
  return t('PNG keeps every pixel.');
}

/** The Export tab's formats, with the names a person knows them by. */
export const EXPORT_FORMATS = [
  ['png', 'PNG'],
  ['jpeg', 'JPEG'],
  ['webp', 'WebP'],
  ['bmp', 'BMP'],
  ['tiff', 'TIFF'],
  ['ico', 'ICO'],
];

/* ── SVG: drawn at a size a person chooses ── */

/** The widest side Image draws an SVG at. */
export const SVG_MOST = 16384;

const UNITS = { px: 1, pt: 4 / 3, pc: 16, mm: 96 / 25.4, cm: 96 / 2.54, in: 96, em: 16, ex: 8 };

/** A length of the SVG's root in pixels, or 0 for none (or a percentage, which has no size alone). */
function lengthOf(value) {
  const m = /^\s*([0-9.]+(?:e[-+]?\d+)?)\s*([a-z%]*)\s*$/i.exec(value || '');
  if (!m || m[2] === '%') return 0;
  const n = Number(m[1]) * (UNITS[m[2].toLowerCase() || 'px'] || 0);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

/**
 * The size an SVG draws at by itself: its root's width and height, or its
 * view box, or the 300 × 150 a browser gives a picture with neither.
 */
export function svgSize(text) {
  const root = /<svg\b([^>]*)>/i.exec(String(text || ''))?.[1] || '';
  const attr = (name) => new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i').exec(root)?.slice(1).find((v) => v !== undefined);
  const box = (attr('viewBox') || '').trim().split(/[\s,]+/).map(Number);
  const bw = box.length === 4 && box[2] > 0 ? box[2] : 0;
  const bh = box.length === 4 && box[3] > 0 ? box[3] : 0;
  let width = lengthOf(attr('width'));
  let height = lengthOf(attr('height'));
  if (width && !height && bw && bh) height = (width * bh) / bw;
  if (height && !width && bw && bh) width = (height * bw) / bh;
  if (!width || !height) {
    width = bw || 300;
    height = bh || 150;
  }
  return { width: Math.max(1, Math.round(width)), height: Math.max(1, Math.round(height)) };
}

/** An SVG drawn at width × height on a clear canvas, as PNG bytes. */
export async function rasterSvg(img, width, height) {
  const k = Math.min(1, SVG_MOST / Math.max(width, height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(width * k));
  canvas.height = Math.max(1, Math.round(height * k));
  const g = canvas.getContext('2d');
  g.imageSmoothingEnabled = true;
  g.imageSmoothingQuality = 'high';
  g.drawImage(img, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  if (!blob) throw new Error(t('The picture could not be drawn at that size.'));
  return blobBytes(blob);
}
