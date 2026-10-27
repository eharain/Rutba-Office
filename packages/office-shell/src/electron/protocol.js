// The rutba:// scheme.
//
// Four hosts, one job each:
//   rutba://app/…            the renderer bundle, served from disk
//   rutba://file/<b64url>    a local file, with byte-range support; a picture Chromium
//                            cannot draw (a TIFF, say) served as a PNG of one page
//                            (?page=N, ?fit=px), and ?info says its pages
//   rutba://thumb/<b64url>   a small JPEG of a local file, made once and kept (thumbs.js)
//   rutba://blob/<id>        bytes the main process is holding for one window
//
// The file host is why video works. A <video src="rutba://file/…"> asks for
// ranges; we answer them, so Chromium seeks a 2 GB recording without anything
// being read into JavaScript. The same host serves images and fonts.
//
// The scheme is registered as standard + secure so the renderer gets a real
// origin (localStorage, fetch, workers) instead of file:// null-origin rules.

import { protocol, net } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Worker } from 'node:worker_threads';
import { fileURLToPath } from 'node:url';
import { hold, heldBlob, release, releaseOwner } from './blobs.js';

export const SCHEME = 'rutba';

export function registerSchemePrivileges() {
  protocol.registerSchemesAsPrivileged([
    {
      scheme: SCHEME,
      privileges: {
        standard: true,
        secure: true,
        supportFetchAPI: true,
        corsEnabled: true,
        stream: true,
        codeCache: true,
      },
    },
  ]);
}

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.avif': 'image/avif',
  '.bmp': 'image/bmp',
  '.ico': 'image/x-icon',
  '.tif': 'image/tiff',
  '.tiff': 'image/tiff',
  '.cur': 'image/x-icon',
  '.psd': 'image/vnd.adobe.photoshop',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.mp4': 'video/mp4',
  '.m4v': 'video/mp4',
  '.webm': 'video/webm',
  '.mkv': 'video/x-matroska',
  '.mov': 'video/quicktime',
  '.ogv': 'video/ogg',
  '.mp3': 'audio/mpeg',
  '.m4a': 'audio/mp4',
  '.wav': 'audio/wav',
  '.flac': 'audio/flac',
  '.ogg': 'audio/ogg',
  '.opus': 'audio/opus',
  '.pdf': 'application/pdf',
  '.txt': 'text/plain; charset=utf-8',
  '.wasm': 'application/wasm',
};

export function mimeFor(p) {
  return MIME[path.extname(p).toLowerCase()] || 'application/octet-stream';
}

export function encodePath(p) {
  return Buffer.from(p, 'utf8').toString('base64url');
}

export function decodePath(token) {
  return Buffer.from(token, 'base64url').toString('utf8');
}

/** A local file URL the renderer can put straight into src=. */
export function fileUrl(p) {
  return `${SCHEME}://file/${encodePath(p)}`;
}

/** The thumbnail of a local file, for a grid tile. */
export function thumbUrl(p, size = 256) {
  return `${SCHEME}://thumb/${encodePath(p)}?s=${size}`;
}

/**
 * Hold bytes for the renderer to fetch by URL — attachments, generated
 * previews. Owned by the window whose request asked, and let go when it
 * closes (blobs.js); `group` lets go of that window's earlier blob of the
 * same group, as the next message read replaces the last one's pictures.
 */
export function holdBlob(bytes, type = 'application/octet-stream', name = '', options = {}) {
  const { id, size } = hold(bytes, type, name, options);
  return { id, url: blobUrl(id), size };
}

/** Where a held blob is fetched: what a thread that mints its own ids writes in its models. */
export function blobUrl(id) {
  return `${SCHEME}://blob/${id}`;
}

export function releaseBlob(id) {
  release(id);
}

export { releaseOwner as releaseBlobsOf };

// The renderer's page is the app host and a file is the file host: two
// origins of one scheme. A window may read what it draws — a frame of a
// clip drawn onto a canvas, say — only if the file host says so, and it
// says so: the window can already read any file through the bridge.
const CORS = { 'access-control-allow-origin': '*', 'access-control-expose-headers': 'content-length, content-range, accept-ranges' };

function rangeResponse(filePath, rangeHeader, type) {
  const size = fs.statSync(filePath).size;
  const m = /^bytes=(\d*)-(\d*)$/.exec(rangeHeader || '');
  if (!m) {
    return new Response(fs.createReadStream(filePath), {
      status: 200,
      headers: { 'content-type': type, 'content-length': String(size), 'accept-ranges': 'bytes', ...CORS },
    });
  }
  let start = m[1] === '' ? null : Number(m[1]);
  let end = m[2] === '' ? null : Number(m[2]);
  if (start === null) {
    // suffix range: the last N bytes
    start = Math.max(0, size - (end ?? 0));
    end = size - 1;
  } else if (end === null || end >= size) {
    end = size - 1;
  }
  if (start > end || start >= size) {
    return new Response(null, { status: 416, headers: { 'content-range': `bytes */${size}`, ...CORS } });
  }
  return new Response(fs.createReadStream(filePath, { start, end }), {
    status: 206,
    headers: {
      'content-type': type,
      'content-length': String(end - start + 1),
      'content-range': `bytes ${start}-${end}/${size}`,
      'accept-ranges': 'bytes',
      ...CORS,
    },
  });
}

/* ── pictures Chromium cannot draw ──────────────────────────────────────── */

// decodePicture (@rutba/imaging/decode) reads them on a thread of its own
// (picture-worker.js), so a large TIFF never holds the windows still; a page
// made is kept a while, as a browser keeps an image, so stepping back and
// forth through a fax's pages does not decode each again.

let imaging = null;
const imagingNow = () => (imaging ||= import('@rutba/imaging/decode'));

/** Whether a file is one a window cannot draw itself (by its name), so it is served decoded. */
async function decodable(target) {
  const { needsDecoding } = await imagingNow();
  return needsDecoding(target);
}

let worker = null;
let nextId = 1;
const waiting = new Map();
const DECODE_MS = 60000;

function pictureWorker() {
  if (worker) return worker;
  const w = new Worker(fileURLToPath(new URL('./picture-worker.js', import.meta.url)));
  w.unref();
  w.on('message', (m) => {
    const job = waiting.get(m.id);
    if (!job) return;
    waiting.delete(m.id);
    clearTimeout(job.timer);
    job.resolve(m);
  });
  const fail = (why) => {
    if (worker === w) worker = null;
    for (const [id, job] of waiting) {
      if (job.worker !== w) continue;
      clearTimeout(job.timer);
      waiting.delete(id);
      job.resolve({ error: `This picture could not be read: ${why}.` });
    }
  };
  w.on('error', (err) => fail(err?.message || 'its reader failed'));
  w.on('exit', () => fail('its reader stopped'));
  worker = w;
  return w;
}

/** Ask the picture thread; a file that takes longer than a minute stops it, and is said so. */
function askWorker(job) {
  return new Promise((resolve) => {
    const id = nextId++;
    const w = pictureWorker();
    const timer = setTimeout(() => {
      waiting.delete(id);
      if (worker === w) worker = null;
      w.terminate().catch(() => {});
      resolve({ error: 'This picture took too long to read, and was stopped.' });
    }, DECODE_MS);
    waiting.set(id, { resolve, timer, worker: w });
    w.postMessage({ id, ...job });
  });
}

/** A JPEG's pixels from Chromium's own decoder (nativeImage), for the JPEG strips inside a TIFF. */
async function jpegDecoder() {
  const { nativeImage } = await import('electron');
  return (bytes) => {
    const img = nativeImage.createFromBuffer(Buffer.from(bytes));
    if (img.isEmpty()) return null;
    const { width, height } = img.getSize();
    const bgra = img.toBitmap();
    const rgba = new Uint8Array(width * height * 4);
    for (let i = 0; i < rgba.length && i + 3 < bgra.length; i += 4) {
      rgba[i] = bgra[i + 2];
      rgba[i + 1] = bgra[i + 1];
      rgba[i + 2] = bgra[i];
      rgba[i + 3] = 255;
    }
    return { width, height, rgba };
  };
}

const made = new Map();
let madeBytes = 0;
const MADE_MAX = 192 * 1024 * 1024;

function keep(key, value) {
  made.set(key, value);
  madeBytes += value.png?.length || 0;
  for (const [k, v] of made) {
    if (madeBytes <= MADE_MAX || k === key) break;
    made.delete(k);
    madeBytes -= v.png?.length || 0;
  }
}

/** A page of a picture as a PNG, or its pages and size (`info`); `{ error }` with the sentence where it cannot be read. */
export async function decodedPicture(target, { info = false, page = 0, fit = 0 } = {}) {
  const stat = await fs.promises.stat(target);
  const key = `${target}|${stat.size}|${Math.round(stat.mtimeMs)}|${info ? 'info' : `${page}|${fit}`}`;
  if (made.has(key)) {
    const hit = made.get(key);
    made.delete(key);
    made.set(key, hit);
    return hit;
  }
  let answer = await askWorker({ file: target, page, fit, info });
  if (answer.needsJpeg) {
    // On this thread, with Chromium's JPEG decoder lent: a JPEG-compressed TIFF.
    const { decodePicture } = await imagingNow();
    try {
      const pic = decodePicture(await fs.promises.readFile(target), path.basename(target), { page, fit, png: true, level: 3, jpeg: await jpegDecoder() });
      answer = { png: pic.png, width: pic.width, height: pic.height, pages: pic.pages, page: pic.page };
    } catch (err) {
      answer = { error: err?.message || String(err) };
    }
  }
  if (!answer.error) keep(key, answer);
  return answer;
}

async function pictureResponse(target, opts) {
  let answer;
  try {
    answer = await decodedPicture(target, opts);
  } catch (err) {
    answer = { error: err?.message || String(err) };
  }
  if (opts.info) {
    const body = answer.error ? { error: answer.error } : answer.info;
    return new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json; charset=utf-8', ...CORS } });
  }
  if (answer.error) {
    return new Response(answer.error, { status: 415, headers: { 'content-type': 'text/plain; charset=utf-8', ...CORS } });
  }
  return new Response(answer.png, {
    status: 200,
    headers: {
      'content-type': 'image/png',
      'content-length': String(answer.png.length),
      'x-rutba-pages': String(answer.pages),
      'x-rutba-page': String(answer.page),
      ...CORS,
      'access-control-expose-headers': 'content-length, x-rutba-pages, x-rutba-page',
    },
  });
}

const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob: rutba:",
  "media-src 'self' blob: rutba:",
  "font-src 'self' data: rutba:",
  "connect-src 'self' data: blob: rutba:",
  "object-src 'self' blob: rutba:",
  "frame-src 'self' blob: rutba:",
  "base-uri 'none'",
  "form-action 'none'",
].join('; ');

/**
 * Automate → scripts: the worker a script runs in may work the script out
 * (`unsafe-eval`) and reach nothing else — no network, no file, no frame.
 */
const WORKER_CSP = "default-src 'none'; script-src 'self' 'unsafe-eval'";

/**
 * The PDF app's page drawer (pdf.js's worker): it compiles its image
 * decoders from WebAssembly and fetches the fonts a PDF names from the
 * app's own files, and reaches nothing else.
 */
const PDF_WORKER_CSP = "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; connect-src 'self' data: blob:";

/**
 * @param {object} o
 * @param {string} o.rendererDir directory holding index.html and the bundle
 * @param {(p: string) => boolean} [o.allowFile] gate for rutba://file reads
 * @param {object} [o.thumbnailer] the thumbnail maker (thumbs.js); without one the thumb host answers 404
 */
export function installProtocol({ rendererDir, allowFile = () => true, thumbnailer = null, packRoots = () => ({}) }) {
  const root = path.resolve(rendererDir);

  protocol.handle(SCHEME, async (request) => {
    const url = new URL(request.url);
    const host = url.host;

    if (host === 'app') {
      let rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      if (rel === '' || !path.extname(rel)) rel = 'index.html';
      // A loaded app pack's own window code, from its archive (docs/APP-PACKS.md):
      // rutba://app/packs/<id>/<file>. A pack not loaded has nothing here.
      let base = root;
      let fromPack = null;
      // A window chunk a loaded pack carries is served from the pack, at the
      // address the core's bundle imports it by (chunks/<name>), so the
      // chunks it shares with the core are the core's own copies.
      if (rel.startsWith('chunks/')) {
        for (const [id, archive] of Object.entries(packRoots())) {
          if (fs.existsSync(path.join(archive, 'renderer', ...rel.split('/')))) {
            base = path.resolve(archive, 'renderer');
            fromPack = id;
            break;
          }
        }
      } else if (!rel.startsWith('packs/') && !fs.existsSync(path.join(root, ...rel.split('/')))) {
        // A file a pack's window starts by name (Worksheets' scripts worker):
        // the core's copy when it has one, else the loaded pack's.
        for (const [id, archive] of Object.entries(packRoots())) {
          if (fs.existsSync(path.join(archive, 'renderer', ...rel.split('/')))) {
            base = path.resolve(archive, 'renderer');
            fromPack = id;
            break;
          }
        }
      }
      if (rel.startsWith('packs/')) {
        const [, id, ...rest] = rel.split('/');
        const archive = packRoots()[id];
        if (!archive) return new Response('not found', { status: 404 });
        base = path.resolve(archive);
        rel = rest.join('/');
      }
      const full = path.resolve(base, rel);
      // Never serve outside the bundle directory or the pack's archive, whatever the URL claims.
      if (full !== base && !full.startsWith(base + path.sep)) {
        return new Response('forbidden', { status: 403 });
      }
      if (!fs.existsSync(full)) return new Response('not found', { status: 404 });
      const res = await net.fetch(pathToFileURL(full).toString());
      const headers = new Headers();
      headers.set('content-type', mimeFor(full));
      if (fromPack) headers.set('x-rutba-pack', fromPack);
      // The renderer runs no remote code and makes no remote requests. Media and
      // images arrive over rutba://file, which is same-scheme.
      headers.set('content-security-policy', rel === 'scripts-worker.js' ? WORKER_CSP : rel === 'pdfjs/pdf.worker.min.mjs' ? PDF_WORKER_CSP : CSP);
      return new Response(res.body, { status: res.status, headers });
    }

    if (host === 'file') {
      const token = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      let target;
      try {
        target = decodePath(token);
      } catch {
        return new Response('bad path', { status: 400 });
      }
      if (!allowFile(target)) return new Response('forbidden', { status: 403 });
      if (!fs.existsSync(target) || fs.statSync(target).isDirectory()) {
        return new Response('not found', { status: 404 });
      }
      // A picture no window can draw from its bytes: a PNG of the page asked for.
      const q = url.searchParams;
      if (!q.has('raw') && (q.has('info') || q.has('page') || (await decodable(target)))) {
        return pictureResponse(target, { info: q.has('info'), page: Number(q.get('page')) || 0, fit: Number(q.get('fit')) || 0 });
      }
      return rangeResponse(target, request.headers.get('range'), mimeFor(target));
    }

    if (host === 'thumb') {
      if (!thumbnailer) return new Response('no thumbnails', { status: 404, headers: { 'x-rutba-thumb': 'none' } });
      const token = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      let target;
      try {
        target = decodePath(token);
      } catch {
        return new Response('bad path', { status: 400 });
      }
      if (!allowFile(target)) return new Response('forbidden', { status: 403 });
      if (!fs.existsSync(target)) return new Response('not found', { status: 404 });
      // One size, so a file has one thumbnail however many places show it.
      const size = Math.max(64, Math.min(512, Number(url.searchParams.get('s')) || 256));
      return thumbnailer.respond(target, { size });
    }

    if (host === 'blob') {
      const id = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      const held = heldBlob(id);
      if (!held) return new Response('not found', { status: 404 });
      const headers = {
        'content-type': held.type,
        'content-length': String(held.bytes.length),
      };
      if (held.name) {
        headers['content-disposition'] = `inline; filename="${held.name.replace(/"/g, '')}"`;
      }
      return new Response(held.bytes, { status: 200, headers });
    }

    return new Response('not found', { status: 404 });
  });
}
