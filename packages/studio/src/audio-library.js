// Video's audio library (owner, D-2026-11-22-10: "the audio library is a list
// of locally added files or remote playable audios and videos only nothing
// more than that").
//
// It is a list and nothing else: each entry is a sound or video file on this
// computer, or the address of one that plays from the web. An entry goes onto
// the timeline as sound. The list is kept with the person's settings, not in
// a project, so it is there for every project on this computer.
//
// This file is the list's own logic, with no window in it, so a test reads it.

/** The settings key the list is kept under. */
export const AUDIO_LIBRARY_KEY = 'video.audioLibrary';

/** The file endings that play as sound: sound files, and video files for their sound. */
export const SOUND_ENDINGS = ['mp3', 'm4a', 'aac', 'wav', 'flac', 'ogg', 'opus'];
export const VIDEO_ENDINGS = ['mp4', 'm4v', 'mov', 'webm', 'mkv', 'ogv'];
export const PLAYABLE_ENDINGS = [...SOUND_ENDINGS, ...VIDEO_ENDINGS];

/** A path's or an address's ending, in small letters, without its dot. */
export function endingOf(name) {
    const clean = String(name || '').split(/[?#]/)[0];
    const m = /\.([A-Za-z0-9]{1,5})$/.exec(clean);
    return m ? m[1].toLowerCase() : '';
}

/** What a file or an address plays as by its ending: 'audio', 'video' or null. */
export function playableKind(name) {
    const e = endingOf(name);
    if (SOUND_ENDINGS.includes(e)) return 'audio';
    if (VIDEO_ENDINGS.includes(e)) return 'video';
    return null;
}

/** A file's name without its folders. */
export const baseName = (p) => String(p || '').split(/[\\/]/).filter(Boolean).pop() || '';

/**
 * Why an address cannot be an entry, or null: only https addresses are
 * fetched, as with a web page used as a picture source. `allowLoopback` is
 * for the checks' own local server.
 */
export function refuseAddress(raw, { allowLoopback = false } = {}) {
    let u;
    try { u = new URL(String(raw || '').trim()); } catch { return 'not-an-address'; }
    if (u.protocol === 'https:') return null;
    if (allowLoopback && u.protocol === 'http:' && u.hostname === '127.0.0.1') return null;
    return 'not-https';
}

/** The name an address is listed by: its file's name, else its host. */
export function addressName(raw) {
    try {
        const u = new URL(String(raw).trim());
        let last = u.pathname.split('/').filter(Boolean).pop() || '';
        try { last = decodeURIComponent(last); } catch { /* kept as it is written */ }
        return last || u.hostname;
    } catch {
        return String(raw || '');
    }
}

const sameFile = (a, b) => String(a).replace(/\\/g, '/').toLowerCase() === String(b).replace(/\\/g, '/').toLowerCase();

/** The list as it was kept, with anything that is not an entry left out. */
export function normaliseLibrary(raw) {
    const out = [];
    const seen = new Set();
    for (const e of Array.isArray(raw) ? raw : []) {
        if (!e || typeof e !== 'object' || !e.id || seen.has(e.id)) continue;
        if (e.kind === 'file' && typeof e.path === 'string' && e.path) {
            out.push({ id: String(e.id), kind: 'file', path: e.path, name: String(e.name || baseName(e.path)) });
        } else if (e.kind === 'remote' && typeof e.url === 'string' && !refuseAddress(e.url, { allowLoopback: true })) {
            out.push({ id: String(e.id), kind: 'remote', url: e.url, name: String(e.name || addressName(e.url)), ...(typeof e.file === 'string' && e.file ? { file: e.file } : {}) });
        } else continue;
        seen.add(String(e.id));
    }
    return out;
}

/** An id no entry of the list has. */
function freshId(list) {
    const taken = new Set(list.map((e) => e.id));
    let n = list.length + 1;
    while (taken.has(`a${n}`)) n += 1;
    return `a${n}`;
}

/**
 * Files from this computer added to the list. A file already listed is not
 * listed twice, and a file that is neither sound nor video is left out.
 * Answers { list, added: the new entries, skipped: the paths left out }.
 */
export function addFileEntries(list, paths) {
    let out = [...list];
    const added = [];
    const skipped = [];
    for (const p of Array.isArray(paths) ? paths : []) {
        if (!p || !playableKind(p)) { if (p) skipped.push(p); continue; }
        if (out.some((e) => e.kind === 'file' && sameFile(e.path, p))) continue;
        const entry = { id: freshId(out), kind: 'file', path: String(p), name: baseName(p) };
        out = [...out, entry];
        added.push(entry);
    }
    return { list: out, added, skipped };
}

/**
 * An address added to the list. Answers { list, entry }, or { list, error }
 * with 'not-an-address', 'not-https' or 'listed' and the list as it was.
 */
export function addRemoteEntry(list, address, { allowLoopback = false } = {}) {
    const url = String(address || '').trim();
    const error = refuseAddress(url, { allowLoopback });
    if (error) return { list, error };
    if (list.some((e) => e.kind === 'remote' && e.url === url)) return { list, error: 'listed' };
    const entry = { id: freshId(list), kind: 'remote', url, name: addressName(url) };
    return { list: [...list, entry], entry };
}

/** The list without one entry. Nothing on any timeline changes: a clip keeps its file. */
export const removeEntry = (list, id) => list.filter((e) => e.id !== id);

/** The list with the file a remote entry was fetched into remembered, so it is fetched once. */
export const rememberFile = (list, id, file) => list.map((e) => (e.id === id && e.kind === 'remote' ? { ...e, file } : e));

/** The entries whose name holds `query`, any case. */
export function searchEntries(list, query) {
    const q = String(query || '').trim().toLowerCase();
    return q ? list.filter((e) => e.name.toLowerCase().includes(q)) : list;
}
