/**
 * The Image app's project file, `.rimg` (decision D2 in docs/STUDIO.md): a
 * zip of the layered picture with everything it is made of inside it, so a
 * project is one file that can be moved, mailed and opened again whole.
 *
 *   mimetype            application/vnd.rutba.image-project+zip, first and
 *                       stored, so the file says what it is from its first
 *                       bytes (as ODF and .rvid do)
 *   doc.json            the project: the studio document (layers, canvas
 *                       size) and the quick edits made to the picture opened
 *   source/<name>       the picture as it was opened, untouched, so the quick
 *                       edits stay editable
 *   bitmaps/<name>.png  the painted bitmaps: the opened picture with its quick
 *                       edits, and every picture added as a layer
 *
 * A layer names a bitmap inside the file by a url `rimg:<path>`; `readRimg`
 * answers the files by those paths and names any a layer asks for that is
 * not there. Unlike .rvid, nothing is referenced outside the file.
 *
 * Pure: bytes in, bytes out. Runs in a page and in Node.
 */

import { zipStored, unzip } from './rvid.js';

export const RIMG_MIME = 'application/vnd.rutba.image-project+zip';
export const RIMG_VERSION = 1;
export const RIMG_SCHEME = 'rimg:';

const utf8 = (s) => new TextEncoder().encode(s);

/** The path inside the file a `rimg:` url names, or null for any other url. */
export const rimgPath = (url) => (typeof url === 'string' && url.startsWith(RIMG_SCHEME) ? url.slice(RIMG_SCHEME.length) : null);

/** A file name safe to keep inside the zip: no folders, no oddities. */
export function safeName(name, fallback = 'picture') {
    const base = String(name || '').split(/[\\/]/).pop().replace(/[^\p{L}\p{N}._ -]+/gu, '_').trim();
    return base.replace(/^\.+/, '') || fallback;
}

/**
 * A project as .rimg bytes. `files` maps a path inside the zip to its bytes
 * (`source/…`, `bitmaps/…`); only those under source/ and bitmaps/ are kept.
 */
export function writeRimg(doc, files = {}) {
    const body = { rimg: RIMG_VERSION, ...doc };
    const entries = [
        { name: 'mimetype', bytes: utf8(RIMG_MIME) },
        { name: 'doc.json', bytes: utf8(JSON.stringify(body, null, 1)) },
    ];
    for (const [name, bytes] of Object.entries(files)) {
        if (!/^(source|bitmaps)\/[^/]+$/.test(name) || !bytes) continue;
        entries.push({ name, bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes) });
    }
    return zipStored(entries);
}

/** The `rimg:` paths a project's layers ask for. */
export function rimgPaths(doc) {
    const out = new Set();
    for (const p of doc?.patches || []) {
        const at = rimgPath(p?.url);
        if (at) out.add(at);
    }
    if (doc?.quick?.source) out.add(doc.quick.source);
    return [...out];
}

/**
 * Read .rimg bytes. Answers { doc, files: { path: Uint8Array }, missing:
 * [path] } where `missing` lists what the layers name and the file lacks.
 */
export async function readRimg(bytes) {
    let parts;
    try {
        parts = await unzip(bytes);
    } catch {
        throw new Error('This is not a Rutba Image project: it is not a zip.');
    }
    const mime = parts.mimetype ? new TextDecoder().decode(parts.mimetype).trim() : null;
    if (mime && mime !== RIMG_MIME) throw new Error('This is not a Rutba Image project.');
    if (!parts['doc.json']) throw new Error('This project is damaged: it has no document in it.');
    let doc;
    try {
        doc = JSON.parse(new TextDecoder().decode(parts['doc.json']));
    } catch {
        throw new Error('This project is damaged: its document cannot be read.');
    }
    if (!doc || typeof doc !== 'object' || !Array.isArray(doc.patches)) throw new Error('This project is damaged: its document has no list of layers.');
    if (Number(doc.rimg) > RIMG_VERSION) throw new Error('This project was made by a newer Rutba Office. Update to open it.');
    delete doc.rimg;
    const files = {};
    for (const [name, data] of Object.entries(parts)) {
        if (/^(source|bitmaps)\//.test(name)) files[name] = data;
    }
    const missing = rimgPaths(doc).filter((p) => !files[p]);
    return { doc, files, missing };
}
