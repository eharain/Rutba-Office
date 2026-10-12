/**
 * The Video app's project file, `.rvid` (decision D2 in docs/STUDIO.md): a
 * zip of the project document and the references to its sources.
 *
 *   mimetype      application/vnd.rutba.video-project+zip, first and stored,
 *                 so the file says what it is from its first bytes (as ODF does)
 *   doc.json      the project: the studio document, clips naming their
 *                 sources by the url (a path, on the desktop) they were added from
 *   sources.json  each source: its url, name, size, length and kind, and its
 *                 path relative to the project's folder
 *
 * The sources themselves are not inside; the project points at them. A
 * project moved together with its media (the relative paths still hold) is
 * found again on open: a source missing at its own path is looked for at its
 * relative one, and the clips are pointed there.
 *
 * Pure: bytes in, bytes out, with the one question it cannot answer itself
 * (whether a file exists) asked of the caller. Runs in a page and in Node.
 */

export const RVID_MIME = 'application/vnd.rutba.video-project+zip';
export const RVID_VERSION = 1;

// ── a small zip, stored entries, read with deflate where the platform has it ──

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

export function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
}

const utf8 = (s) => new TextEncoder().encode(s);

/** A zip of stored entries: [{ name, bytes }] -> Uint8Array. */
export function zipStored(entries) {
    const parts = [];
    const central = [];
    let offset = 0;
    for (const { name, bytes } of entries) {
        const nameBytes = utf8(name);
        const crc = crc32(bytes);
        const local = new Uint8Array(30 + nameBytes.length);
        const lv = new DataView(local.buffer);
        lv.setUint32(0, 0x04034b50, true);
        lv.setUint16(4, 20, true);
        lv.setUint16(6, 0x0800, true); // names in UTF-8
        lv.setUint16(8, 0, true); // stored
        lv.setUint32(14, crc, true);
        lv.setUint32(18, bytes.length, true);
        lv.setUint32(22, bytes.length, true);
        lv.setUint16(26, nameBytes.length, true);
        local.set(nameBytes, 30);
        parts.push(local, bytes);
        const head = new Uint8Array(46 + nameBytes.length);
        const cv = new DataView(head.buffer);
        cv.setUint32(0, 0x02014b50, true);
        cv.setUint16(4, 20, true);
        cv.setUint16(6, 20, true);
        cv.setUint16(8, 0x0800, true);
        cv.setUint16(10, 0, true);
        cv.setUint32(16, crc, true);
        cv.setUint32(20, bytes.length, true);
        cv.setUint32(24, bytes.length, true);
        cv.setUint16(28, nameBytes.length, true);
        cv.setUint32(42, offset, true);
        head.set(nameBytes, 46);
        central.push(head);
        offset += local.length + bytes.length;
    }
    const centralSize = central.reduce((n, c) => n + c.length, 0);
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    const out = new Uint8Array(offset + centralSize + 22);
    let at = 0;
    for (const p of [...parts, ...central, end]) { out.set(p, at); at += p.length; }
    return out;
}

async function inflateRaw(bytes) {
    if (typeof DecompressionStream !== 'function') throw new Error('This project is compressed in a way this system cannot read.');
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** The entries of a zip, by name: { name: Uint8Array }. Stored and deflated entries. */
export async function unzip(input) {
    const b = input instanceof Uint8Array ? input : new Uint8Array(input);
    const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let eocd = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 22 - 65535); i--) {
        if (v.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This is not a Rutba Video project: it is not a zip.');
    const count = v.getUint16(eocd + 10, true);
    let at = v.getUint32(eocd + 16, true);
    const out = {};
    for (let n = 0; n < count; n++) {
        if (v.getUint32(at, true) !== 0x02014b50) throw new Error('This project is damaged: its list of parts is broken.');
        const method = v.getUint16(at + 10, true);
        const compressedSize = v.getUint32(at + 20, true);
        const nameLen = v.getUint16(at + 28, true);
        const extraLen = v.getUint16(at + 30, true);
        const commentLen = v.getUint16(at + 32, true);
        const localAt = v.getUint32(at + 42, true);
        const name = new TextDecoder().decode(b.subarray(at + 46, at + 46 + nameLen));
        const dataAt = localAt + 30 + v.getUint16(localAt + 26, true) + v.getUint16(localAt + 28, true);
        const raw = b.subarray(dataAt, dataAt + compressedSize);
        if (method === 0) out[name] = raw.slice();
        else if (method === 8) out[name] = await inflateRaw(raw);
        at += 46 + nameLen + extraLen + commentLen;
    }
    return out;
}

// ── paths, without a path module ─────────────────────────────────────────────

const sepOf = (p) => (String(p).includes('\\') && !String(p).includes('/') ? '\\' : '/');
const partsOf = (p) => String(p).replace(/\\/g, '/').split('/').filter((x, i) => x !== '' || i === 0);
const isAbsolute = (p) => /^([A-Za-z]:[\\/]|[\\/])/.test(String(p || ''));
const sameDrive = (a, b) => {
    const da = /^([A-Za-z]):/.exec(a)?.[1]?.toLowerCase() || '';
    const db = /^([A-Za-z]):/.exec(b)?.[1]?.toLowerCase() || '';
    return da === db;
};

/** The folder a file is in. */
export function folderOf(p) {
    const parts = partsOf(p);
    parts.pop();
    return parts.join(sepOf(p));
}

/** `target` written relative to `from` (a folder), or null across drives. */
export function relativePath(from, target) {
    if (!from || !target || !isAbsolute(from) || !isAbsolute(target) || !sameDrive(from, target)) return null;
    const windows = /^[A-Za-z]:/.test(from);
    const norm = (x) => (windows ? x.toLowerCase() : x);
    const a = partsOf(from);
    const b = partsOf(target);
    let i = 0;
    while (i < a.length && i < b.length && norm(a[i]) === norm(b[i])) i++;
    return [...a.slice(i).map(() => '..'), ...b.slice(i)].join('/');
}

/** A relative path resolved against a folder. */
export function resolvePath(from, rel) {
    const out = partsOf(from);
    for (const part of String(rel).split(/[\\/]/)) {
        if (!part || part === '.') continue;
        if (part === '..') out.pop();
        else out.push(part);
    }
    return out.join(sepOf(from));
}

// ── the project file ─────────────────────────────────────────────────────────

/** The urls a project's clips name, with what the project knows of each. */
export function sourceList(doc, projectPath = null) {
    const urls = new Set();
    for (const p of doc?.patches || []) if (p?.url) urls.add(p.url);
    const folder = projectPath ? folderOf(projectPath) : null;
    return [...urls].map((url) => {
        const known = doc?.sources?.[url] || {};
        return {
            url,
            name: known.name || String(url).split(/[\\/]/).pop(),
            kind: known.kind || null,
            size: known.size ?? null,
            duration: known.duration ?? null,
            relative: folder ? relativePath(folder, url) : null,
        };
    });
}

/** A project as .rvid bytes. `projectPath` is where it is being saved, for the relative paths. */
export function writeRvid(doc, { projectPath = null } = {}) {
    const body = { rvid: RVID_VERSION, ...doc };
    return zipStored([
        { name: 'mimetype', bytes: utf8(RVID_MIME) },
        { name: 'doc.json', bytes: utf8(JSON.stringify(body, null, 1)) },
        { name: 'sources.json', bytes: utf8(JSON.stringify({ sources: sourceList(doc, projectPath) }, null, 1)) },
    ]);
}

/**
 * Read .rvid bytes. `exists(path)` answers whether a source is where the
 * project says (sync or async); a source that is not, but is at its path
 * relative to the project's folder, is pointed there. Answers { doc,
 * sources, moved: [{ from, to }], missing: [url] }.
 */
export async function readRvid(bytes, { projectPath = null, exists = null } = {}) {
    const parts = await unzip(bytes);
    const mime = parts.mimetype ? new TextDecoder().decode(parts.mimetype).trim() : null;
    if (mime && mime !== RVID_MIME) throw new Error('This is not a Rutba Video project.');
    if (!parts['doc.json']) throw new Error('This project is damaged: it has no document in it.');
    let doc;
    try {
        doc = JSON.parse(new TextDecoder().decode(parts['doc.json']));
    } catch {
        throw new Error('This project is damaged: its document cannot be read.');
    }
    if (!doc || typeof doc !== 'object' || !Array.isArray(doc.patches)) throw new Error('This project is damaged: its document has no clips list.');
    if (Number(doc.rvid) > RVID_VERSION) throw new Error('This project was made by a newer Rutba Office. Update to open it.');
    delete doc.rvid;
    let sources = [];
    try {
        sources = parts['sources.json'] ? JSON.parse(new TextDecoder().decode(parts['sources.json'])).sources || [] : [];
    } catch {
        sources = [];
    }
    const moved = [];
    const missing = [];
    if (exists) {
        const folder = projectPath ? folderOf(projectPath) : null;
        const remap = new Map();
        for (const s of sources) {
            if (!s?.url) continue;
            if (await exists(s.url)) continue;
            const there = folder && s.relative ? resolvePath(folder, s.relative) : null;
            if (there && there !== s.url && await exists(there)) {
                remap.set(s.url, there);
                moved.push({ from: s.url, to: there });
            } else {
                missing.push(s.url);
            }
        }
        if (remap.size) {
            doc.patches = doc.patches.map((p) => (p?.url && remap.has(p.url) ? { ...p, url: remap.get(p.url) } : p));
            const known = { ...(doc.sources || {}) };
            for (const [from, to] of remap) {
                if (known[from]) { known[to] = known[from]; delete known[from]; }
            }
            doc.sources = known;
            sources = sources.map((s) => (remap.has(s.url) ? { ...s, url: remap.get(s.url) } : s));
        }
    }
    return { doc, sources, moved, missing };
}
