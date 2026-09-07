// The rutba:// scheme.
//
// Three hosts, one job each:
//   rutba://app/…            the renderer bundle, served from disk
//   rutba://file/<b64url>    a local file, with byte-range support
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

const blobs = new Map();
let blobSeq = 0;

/** Hold bytes for the renderer to fetch by URL — attachments, generated previews. */
export function holdBlob(bytes, type = 'application/octet-stream', name = '') {
  const id = `b${++blobSeq}`;
  blobs.set(id, { bytes: Buffer.from(bytes), type, name });
  return { id, url: `${SCHEME}://blob/${id}`, size: bytes.length };
}

export function releaseBlob(id) {
  blobs.delete(id);
}

function rangeResponse(filePath, rangeHeader, type) {
  const size = fs.statSync(filePath).size;
  const m = /^bytes=(\d*)-(\d*)$/.exec(rangeHeader || '');
  if (!m) {
    return new Response(fs.createReadStream(filePath), {
      status: 200,
      headers: { 'content-type': type, 'content-length': String(size), 'accept-ranges': 'bytes' },
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
    return new Response(null, { status: 416, headers: { 'content-range': `bytes */${size}` } });
  }
  return new Response(fs.createReadStream(filePath, { start, end }), {
    status: 206,
    headers: {
      'content-type': type,
      'content-length': String(end - start + 1),
      'content-range': `bytes ${start}-${end}/${size}`,
      'accept-ranges': 'bytes',
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
 * @param {object} o
 * @param {string} o.rendererDir directory holding index.html and the bundle
 * @param {(p: string) => boolean} [o.allowFile] gate for rutba://file reads
 */
export function installProtocol({ rendererDir, allowFile = () => true }) {
  const root = path.resolve(rendererDir);

  protocol.handle(SCHEME, async (request) => {
    const url = new URL(request.url);
    const host = url.host;

    if (host === 'app') {
      let rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      if (rel === '' || !path.extname(rel)) rel = 'index.html';
      const full = path.resolve(root, rel);
      // Never serve outside the bundle directory, whatever the URL claims.
      if (full !== root && !full.startsWith(root + path.sep)) {
        return new Response('forbidden', { status: 403 });
      }
      if (!fs.existsSync(full)) return new Response('not found', { status: 404 });
      const res = await net.fetch(pathToFileURL(full).toString());
      const headers = new Headers();
      headers.set('content-type', mimeFor(full));
      // The renderer runs no remote code and makes no remote requests. Media and
      // images arrive over rutba://file, which is same-scheme.
      headers.set('content-security-policy', CSP);
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
      return rangeResponse(target, request.headers.get('range'), mimeFor(target));
    }

    if (host === 'blob') {
      const id = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      const held = blobs.get(id);
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
