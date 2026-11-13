// One add-on file fetched into a staging folder (C4): a `.part` file resumed
// with Range and If-Range after a cut or a pause, its SHA-512 taken while it
// streams, and renamed to its own name only once its size and hash are the
// ones the signed feed gives. Nothing half-written is ever under its own name.
//
// Only addresses of the packages repository's releases are fetched, or files
// under a folder an administrator named (RUTBA_ADDONS_SOURCE); the feed is
// signed, but an address is still checked before it is asked for.

import crypto from 'node:crypto';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

export const RELEASES = 'https://github.com/eharain/rutba-office-packages/releases/download/';
export const releaseUrl = (version) => `${RELEASES}v${version}/`;

/** Whether an address may be fetched: the packages releases, or a file under the named source folder. */
export function allowedUrl(url, { source = null } = {}) {
  const u = String(url || '');
  if (u.startsWith(RELEASES) && !u.slice(RELEASES.length).split('/').some((s) => s === '..' || s === '')) return true;
  if (source && u.startsWith('file:')) {
    try {
      const p = path.resolve(fileURLToPath(u));
      const r = path.relative(path.resolve(source), p);
      return r !== '' && !r.startsWith('..') && !path.isAbsolute(r);
    } catch {
      return false;
    }
  }
  return false;
}

/**
 * Where a feed's file is fetched from with a source folder named: the same
 * release-relative path under that folder (v<version>/<asset>), so a share
 * can hold the release's files as they are.
 */
export function fromSource(url, source) {
  if (!source || !String(url).startsWith(RELEASES)) return url;
  return pathToFileURL(path.join(source, ...String(url).slice(RELEASES.length).split('/'))).href;
}

/** Bytes free on the disk that holds a folder, or null when it cannot be told. */
export function freeSpace(dir) {
  try {
    const s = fs.statfsSync(dir);
    return Number(s.bavail) * Number(s.bsize);
  } catch {
    return null;
  }
}

export class DownloadError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

const hashOfPart = (file) => {
  const hash = crypto.createHash('sha512');
  const fd = fs.openSync(file, 'r');
  try {
    const chunk = Buffer.alloc(1024 * 1024);
    for (let at = 0; ; ) {
      const n = fs.readSync(fd, chunk, 0, chunk.length, at);
      if (n <= 0) break;
      hash.update(chunk.subarray(0, n));
      at += n;
    }
  } finally {
    fs.closeSync(fd);
  }
  return hash;
};

/**
 * Fetches one file to `to`, by way of `to.part` (and `to.part.json`, which
 * keeps the validator If-Range sends).
 * @param {object} o
 * @param {string} o.url
 * @param {number} o.size
 * @param {string} o.sha512        base64
 * @param {string} o.to
 * @param {Function} o.fetchImpl   fetch(url, init) -> Response
 * @param {AbortSignal} [o.signal] a pause or a cancel
 * @param {(n: number) => void} [o.onBytes]  bytes newly written
 * @param {string|null} [o.source]
 */
export async function fetchToFile({ url, size, sha512, to, fetchImpl, signal = null, onBytes = () => {}, source = null }) {
  if (fs.existsSync(to) && fs.statSync(to).size === size) return to;
  const address = fromSource(url, source);
  if (!allowedUrl(address, { source })) throw new DownloadError('address', 'An add-on file is at an address Rutba Office does not fetch from.');
  fs.mkdirSync(path.dirname(to), { recursive: true });
  const part = `${to}.part`;
  const meta = `${part}.json`;

  let hash = null;
  // A file in a named folder: copied, then checked as a download is.
  if (address.startsWith('file:')) {
    const from = fileURLToPath(address);
    if (!fs.existsSync(from)) throw new DownloadError('missing', `${path.basename(from)} is not in the folder of add-ons.`);
    fs.copyFileSync(from, part);
    onBytes(fs.statSync(part).size);
    hash = hashOfPart(part);
  } else {
    let have = fs.existsSync(part) ? fs.statSync(part).size : 0;
    let validator = null;
    try {
      validator = JSON.parse(fs.readFileSync(meta, 'utf8')).validator || null;
    } catch {
      validator = null;
    }
    if (have > size || (have > 0 && !validator)) {
      fs.rmSync(part, { force: true });
      have = 0;
    }
    const headers = { 'user-agent': 'Rutba-Office' };
    if (have > 0) Object.assign(headers, { range: `bytes=${have}-`, 'if-range': validator });
    const response = await fetchImpl(address, { headers, signal, redirect: 'follow' });
    if (!(response.status === 200 || response.status === 206)) throw new DownloadError('http', `The add-on file could not be fetched (${response.status}).`);
    // 200 to a ranged request: the file changed, or ranges are not served. Start again.
    if (response.status === 200 && have > 0) {
      fs.rmSync(part, { force: true });
      have = 0;
    }
    const v = response.headers?.get?.('etag') || response.headers?.get?.('last-modified') || null;
    fs.writeFileSync(meta, JSON.stringify({ url, validator: v }));
    // What is already there is hashed once; the rest as it arrives.
    hash = have > 0 ? hashOfPart(part) : crypto.createHash('sha512');
    const fd = fs.openSync(part, have > 0 ? 'a' : 'w');
    try {
      if (response.body?.getReader) {
        const reader = response.body.getReader();
        for (;;) {
          if (signal?.aborted) {
            reader.cancel().catch(() => {});
            throw new DownloadError('stopped', 'The download was stopped.');
          }
          const { done, value } = await reader.read();
          if (done) break;
          const b = Buffer.from(value);
          if (have + b.length > size) throw new DownloadError('size', 'An add-on file is larger than its list says.');
          fs.writeSync(fd, b);
          hash.update(b);
          have += b.length;
          onBytes(b.length);
        }
      } else {
        const b = Buffer.from(await response.arrayBuffer());
        if (have + b.length > size) throw new DownloadError('size', 'An add-on file is larger than its list says.');
        fs.writeSync(fd, b);
        hash.update(b);
        have += b.length;
        onBytes(b.length);
      }
    } finally {
      fs.closeSync(fd);
    }
  }

  if (fs.statSync(part).size !== size) {
    throw new DownloadError('size', 'An add-on file is not the size its list says.');
  }
  if (hash.digest('base64') !== sha512) {
    fs.rmSync(part, { force: true });
    fs.rmSync(meta, { force: true });
    throw new DownloadError('hash', 'An add-on file is not the one its list names. It was not kept.');
  }
  fs.rmSync(meta, { force: true });
  fs.rmSync(to, { force: true });
  fs.renameSync(part, to);
  return to;
}
