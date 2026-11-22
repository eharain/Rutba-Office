// Sound or video at a web address, fetched into a file (Video's audio library,
// D-2026-11-22-10: the list holds files from the computer and "remote playable
// audios and videos").
//
// A window cannot play or read a web address itself: its content policy lets
// it reach Office's own files only. So an address in the list is fetched here,
// once, into a file the timeline then reads as it reads any other. Only an
// https address is fetched, with no cookies and no referrer, every redirect
// held to the same rule, and only when what comes back is sound or video and
// no larger than the limit. A page that plays a video (a video site's watch
// page) is a web page, not a file, and is refused in plain words. The file
// is marked as from the internet, as a browser marks a download, so any
// other program that opens it knows where it came from.

import nodeFs from 'node:fs';
import nodePath from 'node:path';
import { markFromInternet } from '@rutba/office-shell/runs';

/** Why an address is not fetched, or null: the rule a web page source keeps (web-source.js), said again here so this file needs no window. */
export function refuseAddress(raw, { allowLoopback = false } = {}) {
  let u;
  try { u = new URL(String(raw)); } catch { return 'not-an-address'; }
  if (u.protocol === 'https:') return null;
  if (allowLoopback && u.protocol === 'http:' && u.hostname === '127.0.0.1') return null;
  return 'not-https';
}

/** The largest file fetched: a long video, never a whole disk. */
export const REMOTE_MEDIA_LIMIT = 1024 * 1024 * 1024;

const ENDINGS = ['mp3', 'm4a', 'aac', 'wav', 'flac', 'ogg', 'opus', 'mp4', 'm4v', 'mov', 'webm', 'mkv', 'ogv'];

/** The ending a content type is kept under, when the address has none of its own. */
const TYPE_ENDING = {
  'audio/mpeg': 'mp3', 'audio/mp3': 'mp3', 'audio/mp4': 'm4a', 'audio/x-m4a': 'm4a', 'audio/aac': 'aac', 'audio/wav': 'wav', 'audio/x-wav': 'wav', 'audio/wave': 'wav',
  'audio/flac': 'flac', 'audio/x-flac': 'flac', 'audio/ogg': 'ogg', 'application/ogg': 'ogg', 'audio/opus': 'opus', 'audio/webm': 'webm',
  'video/mp4': 'mp4', 'video/x-m4v': 'm4v', 'video/quicktime': 'mov', 'video/webm': 'webm', 'video/x-matroska': 'mkv', 'video/ogg': 'ogv',
};

const endingOf = (name) => (/\.([A-Za-z0-9]{1,5})$/.exec(String(name || '').split(/[?#]/)[0])?.[1] || '').toLowerCase();

/**
 * The ending the fetched file is kept under, or null when what the address
 * answers with is neither sound nor video: by its content type, else (a
 * server that says only "a file") by the address's own ending.
 */
export function mediaEnding(address, contentType) {
  const type = String(contentType || '').split(';')[0].trim().toLowerCase();
  const named = endingOf(new URL(address).pathname);
  const known = ENDINGS.includes(named) ? named : null;
  if (TYPE_ENDING[type]) return known || TYPE_ENDING[type];
  if (/^(audio|video)\//.test(type)) return known;
  if (!type || type === 'application/octet-stream' || type === 'binary/octet-stream') return known;
  return null;
}

/** A file name for the fetched file: the address's own, made safe for a folder, under `ending`. */
export function mediaFileName(address, ending) {
  const u = new URL(address);
  let last = u.pathname.split('/').filter(Boolean).pop() || '';
  try { last = decodeURIComponent(last); } catch { /* kept as it is written */ }
  const stem = (last ? last.replace(/\.[A-Za-z0-9]{1,5}$/, '') : u.hostname).replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').replace(/[. ]+$/, '').trim().slice(0, 80) || u.hostname;
  return `${stem}.${ending}`;
}

const say = (code) => ({
  'not-an-address': 'That is not a web address.',
  'not-https': 'A sound or video on the web must start https://.',
}[code] || code);

/**
 * @param {{ allowLoopback?: boolean, fetch?: typeof fetch, fs?: typeof nodeFs, limit?: number, platform?: string }} [o]
 *   `allowLoopback` lets a check run fetch from its own local server.
 */
export function createRemoteMedia({ allowLoopback = false, fetch: fetchIt = globalThis.fetch, fs = nodeFs, limit = REMOTE_MEDIA_LIMIT, platform = process.platform } = {}) {
  /** The answer at an address, every redirect on the way held to https. */
  async function reach(address) {
    let at = String(address || '').trim();
    for (let hop = 0; hop < 6; hop++) {
      const refused = refuseAddress(at, { allowLoopback });
      if (refused) throw new Error(say(refused));
      const res = await fetchIt(at, { redirect: 'manual', referrerPolicy: 'no-referrer', credentials: 'omit' });
      if (res.status >= 300 && res.status < 400 && res.headers.get('location')) {
        at = new URL(res.headers.get('location'), at).href;
        try { await res.body?.cancel?.(); } catch { /* nothing left to let go of */ }
        continue;
      }
      return { res, at };
    }
    throw new Error('That address sends the request on too many times.');
  }

  /**
   * Fetch the sound or video at `url` into folder `dir`.
   * Answers { path, name, size, type }; a file already there under the same
   * name and of the same size is kept and answered as it is.
   */
  async function webMedia({ url, dir } = {}) {
    if (!dir || !nodePath.isAbsolute(String(dir))) throw new Error('There is no folder to keep the file in.');
    const { res, at } = await reach(url);
    const drop = async () => { try { await res.body?.cancel?.(); } catch { /* already closed */ } };
    if (!res.ok) { await drop(); throw new Error(`The address answered ${res.status}.`); }
    const type = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
    const ending = mediaEnding(at, type) || mediaEnding(String(url).trim(), type);
    if (!ending) { await drop(); throw new Error('That address is not a sound or a video file. A page that plays one is not a file.'); }
    const stated = Number(res.headers.get('content-length'));
    if (Number.isFinite(stated) && stated > limit) { await drop(); throw new Error('That file is too large to add.'); }
    const name = mediaFileName(String(url).trim(), ending);
    const file = nodePath.join(String(dir), name);
    const there = await fs.promises.stat(file).catch(() => null);
    if (there?.isFile() && Number.isFinite(stated) && stated > 0 && there.size === stated) { await drop(); return { path: file, name, size: there.size, type }; }

    await fs.promises.mkdir(String(dir), { recursive: true });
    const part = `${file}.part`;
    const out = await fs.promises.open(part, 'w');
    let size = 0;
    try {
      if (!res.body) throw new Error('The address sent nothing.');
      for await (const chunk of res.body) {
        size += chunk.byteLength;
        if (size > limit) throw new Error('That file is too large to add.');
        await out.write(chunk);
      }
      if (!size) throw new Error('The address sent nothing.');
      await out.close();
      await fs.promises.rename(part, file);
      // Best effort: a disk that holds no such mark still keeps the file.
      await markFromInternet(file, { write: fs.promises.writeFile, platform });
    } catch (err) {
      await out.close().catch(() => {});
      await drop();
      await fs.promises.rm(part, { force: true }).catch(() => {});
      throw err;
    }
    return { path: file, name, size, type };
  }

  return { webMedia };
}
