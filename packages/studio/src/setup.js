// A setup (owner, 2026-11-05, D-2026-11-05-13 and -14): the sources and how
// each is placed, the scenes, the sounds and their levels, Back soon, the shots
// Auto follows, the picture size, frame rate, quality and save folder, and in
// Live Broadcast which streams are on. Any number of them, shared by the
// Recorder and Live Broadcast; each app opens on the one it last used. Saved
// as they change, and as `.rstudio` files, a zip of setup.json and every
// picture and music file the setup uses, so one opens complete on another
// computer. Stream keys never go in one.
//
// What a person has does not travel by this computer's ids: a window's capture
// id changes when its program restarts (the "setup lost" fault), so on every
// open each source is found again by what it is (rematch).
//
// Nothing here touches a window or a disk: the store and the file are the
// caller's. So all of it is tested in Node.

import { normaliseStack, stackFromSettings, materialiseScenes, newSetupStack, BACK_SOON } from './stack.js';
import { normaliseStudio } from './compositor.js';
import { zipStored, unzip } from './rvid.js';

export const SETUP_VERSION = 1;
export const SETUP_EXT = 'rstudio';
export const SETUP_MIME = 'application/vnd.rutba.studio-setup+zip';

/** The apps a setup can have been last used in. */
export const SETUP_APPS = ['recorder', 'broadcast'];

/** The settings kept with each setup; the rest are the same whichever setup is open (A4, A6). */
export const SETUP_KEYS = ['size', 'fps', 'quality', 'folder', 'liveCopy', 'isolated'];
export const GLOBAL_KEYS = ['countdown', 'minimise', 'shot', 'shotTimer'];

/** Where the store keeps them: the list, one key for each setup, and the settings that are the same in all. */
export const STORE = Object.freeze({ index: 'studio.setups', setup: (id) => `studio.setup.${id}`, global: 'studio.global' });

const str = (v, fallback = '', max = 200) => (typeof v === 'string' ? v.slice(0, max) : fallback);
const pick = (v, list, fallback) => (list.includes(v) ? v : fallback);
const pickKeys = (o, keys) => Object.fromEntries(keys.map((k) => [k, o[k]]));

/** The settings a setup keeps, from the studio's whole set. */
export const setupSettings = (studio) => pickKeys(normaliseStudio(studio), SETUP_KEYS);

/** The settings every setup shares, from the studio's whole set. */
export const globalSettings = (studio) => pickKeys(normaliseStudio(studio), GLOBAL_KEYS);

/** A setup's settings and the shared ones, as the one set the studio works from. */
export const joinSettings = (settings, global) => normaliseStudio({ ...(settings || {}), ...(global || {}) });

/** One stream as a setup may carry it: its name and address, and never its key. */
function normaliseStream(raw) {
    if (!raw || typeof raw !== 'object' || typeof raw.id !== 'string' || !raw.id) return null;
    return { id: raw.id.slice(0, 80), name: str(raw.name, '', 120), kind: str(raw.kind, '', 40), url: str(raw.url, '', 2000), tall: Boolean(raw.tall) };
}

/** A stored setup, or anything, as a setup the studio can use. */
export function normaliseSetup(raw, { app = 'recorder' } = {}) {
    const r = raw && typeof raw === 'object' ? raw : {};
    return {
        version: SETUP_VERSION,
        id: str(r.id, '', 60),
        name: str(r.name, '', 120),
        app: pick(r.app, SETUP_APPS, app),
        usedAt: Number.isFinite(Number(r.usedAt)) ? Number(r.usedAt) : 0,
        // The AI tools' own: a setup of its own for a recording they ask for, never in the menu (A8).
        tools: Boolean(r.tools),
        stack: normaliseStack(r.stack),
        settings: setupSettings(r.settings),
        // The streams that are on in this setup (A12); what they are is kept per profile, with their keys.
        liveOn: (Array.isArray(r.liveOn) ? r.liveOn : []).filter((x) => typeof x === 'string' && x).slice(0, 200).map((x) => x.slice(0, 80)),
        // In a file only: the streams it was saved with, without their keys.
        streams: (Array.isArray(r.streams) ? r.streams : []).map(normaliseStream).filter(Boolean).slice(0, 200),
    };
}

/** A new setup from a stack (the four quick scenes when there is none). */
export function newSetup({ id, name, app = 'recorder', stack = null, settings = null, now = 0, screen = null, hasCamera = true, names = {} }) {
    return normaliseSetup({ id, name, app, usedAt: now, stack: stack || newSetupStack({ screen, hasCamera, names }), settings: settings || {} }, { app });
}

/** A copy of a setup, under a new id and name. */
export function duplicateSetup(setup, { id, name, now = 0 }) {
    return normaliseSetup({ ...setup, id, name, usedAt: now, tools: false }, { app: setup.app });
}

/** The stream removed from My streams drops out of a setup too (A12). */
export function dropStream(setup, streamId) {
    return { ...setup, liveOn: setup.liveOn.filter((x) => x !== streamId) };
}

/* ── the list ──────────────────────────────────────────────────────────── */

/** The stored list, or anything, as a list. */
export function normaliseIndex(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const seen = new Set();
    const setups = [];
    for (const x of Array.isArray(r.setups) ? r.setups : []) {
        const e = entryOf(x);
        if (!e || seen.has(e.id)) continue;
        seen.add(e.id);
        setups.push(e);
    }
    const current = {};
    for (const app of SETUP_APPS) current[app] = str(r.current?.[app], '', 60);
    return { version: SETUP_VERSION, setups, current };
}

function entryOf(raw) {
    if (!raw || typeof raw !== 'object' || typeof raw.id !== 'string' || !raw.id) return null;
    return { id: raw.id.slice(0, 60), name: str(raw.name, '', 120), app: pick(raw.app, SETUP_APPS, 'recorder'), usedAt: Number.isFinite(Number(raw.usedAt)) ? Number(raw.usedAt) : 0, tools: Boolean(raw.tools) };
}

/** The list with a setup put in, or changed in its place. */
export function withSetup(index, setup) {
    const e = entryOf(setup);
    const i = index.setups.findIndex((x) => x.id === e.id);
    return { ...index, setups: i < 0 ? [...index.setups, e] : index.setups.map((x, n) => (n === i ? e : x)) };
}

/** The setup an app opens on: the one it last used, or the newest used, or none. */
export function openingSetup(index, app) {
    const own = index.current[app];
    if (own && index.setups.some((x) => x.id === own && !x.tools)) return own;
    return menuOrder(index)[0]?.id || null;
}

/** The setups as the menu lists them, newest used first; the AI tools' own is not among them. */
export function menuOrder(index) {
    return index.setups.filter((x) => !x.tools).slice().sort((a, b) => b.usedAt - a.usedAt);
}

/** A setup opened in an app: it is now that app's, and the newest used. */
export function markUsed(index, id, app, now) {
    return { ...index, setups: index.setups.map((x) => (x.id === id ? { ...x, app, usedAt: now } : x)), current: { ...index.current, [app]: id } };
}

/**
 * A setup taken out of the list, and from any app that was on it: that app
 * goes to the newest one left. Answers the list.
 */
export function withoutSetup(index, id) {
    const setups = index.setups.filter((x) => x.id !== id);
    const next = { ...index, setups };
    const newest = menuOrder(next)[0]?.id || '';
    const current = {};
    for (const app of SETUP_APPS) current[app] = index.current[app] === id ? newest : index.current[app];
    return { ...next, current };
}

/** A name not among `taken`: the name itself, or "name 2", "name 3". */
export function uniqueName(name, taken) {
    const low = new Set(taken.map((x) => String(x).toLowerCase()));
    if (!low.has(String(name).toLowerCase())) return name;
    for (let n = 2; ; n++) {
        const next = `${name} ${n}`;
        if (!low.has(next.toLowerCase())) return next;
    }
}

/** A new setup id, not among `taken`. */
export function setupId(taken, now = Date.now()) {
    const used = new Set(taken);
    for (let n = 0; ; n++) {
        const id = `setup-${now.toString(36)}${n ? `-${n}` : ''}`;
        if (!used.has(id)) return id;
    }
}

/* ── undo and redo ─────────────────────────────────────────────────────── */

const HISTORY_MAX = 200;

/** Nothing to undo or redo yet. */
export const emptyHistory = () => ({ back: [], ahead: [] });

/** The stack as it was before a change, kept to undo to. Whatever was undone cannot be redone after a new change. */
export function historyPush(history, before) {
    return { back: [...history.back, before].slice(-HISTORY_MAX), ahead: [] };
}

/** Answers { history, stack } one change back, or null when there is nothing to undo. */
export function historyUndo(history, current) {
    if (!history.back.length) return null;
    const stack = history.back[history.back.length - 1];
    return { history: { back: history.back.slice(0, -1), ahead: [current, ...history.ahead].slice(0, HISTORY_MAX) }, stack };
}

/** Answers { history, stack } one change on, or null when there is nothing to redo. */
export function historyRedo(history, current) {
    if (!history.ahead.length) return null;
    const [stack, ...ahead] = history.ahead;
    return { history: { back: [...history.back, current].slice(-HISTORY_MAX), ahead }, stack };
}

/* ── finding sources again ─────────────────────────────────────────────── */

const EM_DASH = ` ${String.fromCharCode(0x2014)} `;

/** The program a window's title names: what follows its last " - " or em dash. */
export function programOf(title) {
    const t = String(title || '');
    const at = Math.max(t.lastIndexOf(' - '), t.lastIndexOf(EM_DASH));
    return at >= 0 ? t.slice(at + 3).trim() : '';
}

/** How many edits apart two strings are; used to find the closest title. */
function editDistance(a, b) {
    const m = a.length;
    const n = b.length;
    let prev = Array.from({ length: n + 1 }, (_, j) => j);
    for (let i = 1; i <= m; i++) {
        const row = [i];
        for (let j = 1; j <= n; j++) row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        prev = row;
    }
    return prev[n];
}

const sizeOfSource = (s) => (s?.display?.width && s?.display?.height ? { w: s.display.width, h: s.display.height } : null);
const sameSize = (a, b) => Boolean(a && b && a.w === b.w && a.h === b.h);

/** One screen found again: the same id and size; the screen of the same size; or the main screen, saying which it was. */
function matchScreen(it, screens) {
    if (!screens.length) return { how: 'missing' };
    const same = screens.find((s) => s.id === it.ref);
    if (same && (!it.size || !sizeOfSource(same) || sameSize(it.size, sizeOfSource(same)))) return { how: 'same', source: same };
    const bySize = it.size ? screens.find((s) => sameSize(it.size, sizeOfSource(s))) : null;
    if (bySize) return { how: 'size', source: bySize };
    return { how: 'main', source: screens[0], was: it.size ? `${it.name || ''} (${it.size.w} × ${it.size.h})`.trim() : it.name || '' };
}

/** One window found again: the very same; the same title; a window of the same program with the closest title; or none. */
function matchWindow(it, windows, used) {
    const free = windows.filter((s) => !used.has(s.id));
    const same = free.find((s) => s.id === it.ref);
    if (same) return { how: 'same', source: same };
    const titled = free.find((s) => s.name === it.name);
    if (titled) return { how: 'title', source: titled };
    const program = programOf(it.name);
    if (program) {
        const kin = free.filter((s) => programOf(s.name) === program);
        kin.sort((a, b) => editDistance(a.name, it.name) - editDistance(b.name, it.name));
        if (kin[0]) return { how: 'program', source: kin[0] };
    }
    return { how: 'missing' };
}

/** One device found again: the same id; the same name (each device once, so two alike stay two); or none. */
function matchDevice(ref, name, list, used) {
    if (!ref) return { how: list.length ? 'default' : 'missing', source: list[0] || null };
    const same = list.find((d) => d.id === ref);
    if (same) return { how: 'same', source: same };
    const named = list.find((d) => !used.has(d.id) && d.name === name);
    return named ? { how: 'name', source: named } : { how: 'missing' };
}

/**
 * A setup's sources found again by what they are, on this computer as it is
 * now. Answers { setup, found, missing }: `setup` with the refs and names
 * brought up to date; `found` and `missing` each a list of
 * { id, scope: 'item'|'sound', kind, name, how, was }. What is not found is
 * kept where it was, never dropped: it is shown as Not found with Choose….
 * With `devices` null (they cannot be named yet), cameras and microphones
 * are left as they are, and said neither found nor missing.
 */
export function rematch(setup, sources, devices) {
    const screens = (sources || []).filter((s) => s.kind === 'screen');
    const windows = (sources || []).filter((s) => s.kind === 'window');
    const cameras = devices?.cameras || [];
    const mics = devices?.mics || [];
    const trustDevices = devices != null;
    const usedWindows = new Set();
    const usedCameras = new Set();
    const usedMics = new Set();
    const found = [];
    const missing = [];
    const note = (rec, ok) => (ok ? found : missing).push(rec);

    const items = setup.stack.items.map((it) => {
        if (it.kind === 'title' || it.kind === 'picture') return it;
        // A window yet to be chosen, in a new setup, is not a window lost.
        if (it.kind === 'window' && !it.ref && !it.name) return it;
        if (it.kind === 'screen' && !it.ref && !it.name && !screens.length) return it;
        if (it.kind === 'camera' && !trustDevices) return it;
        let m;
        if (it.kind === 'screen') m = matchScreen(it, screens);
        else if (it.kind === 'window') m = matchWindow(it, windows, usedWindows);
        else m = matchDevice(it.ref, it.name, cameras, usedCameras);
        const rec = { id: it.id, scope: 'item', kind: it.kind, name: it.name, how: m.how, was: m.was || '' };
        if (m.how === 'missing') { note(rec, false); return it; }
        note(rec, true);
        if (it.kind === 'window') usedWindows.add(m.source.id);
        if (it.kind === 'camera' && m.source) usedCameras.add(m.source.id);
        const next = { ...it };
        if (it.kind === 'screen') {
            next.ref = m.source.id;
            next.name = m.source.name || it.name;
            // The main screen stands in until one is chosen: it says which it was, and is looked for again next time.
            if (m.how === 'main') next.was = m.was;
            else { next.was = ''; next.size = sizeOfSource(m.source) || it.size; }
        } else if (it.kind === 'window') {
            next.ref = m.source.id;
            next.name = m.source.name || it.name;
        } else if (m.how !== 'default') {
            next.ref = m.source.id;
            next.name = m.source.name || it.name;
        }
        rec.name = next.name;
        return next;
    });

    const sounds = setup.stack.sounds.map((s) => {
        if (s.kind !== 'mic' || !trustDevices) return s;
        const m = matchDevice(s.ref, s.name, mics, usedMics);
        const rec = { id: s.id, scope: 'sound', kind: s.kind, name: s.name, how: m.how, was: '' };
        if (m.how === 'missing') { note(rec, false); return s; }
        note(rec, true);
        if (m.source) usedMics.add(m.source.id);
        if (m.how === 'default') return s;
        rec.name = m.source.name || s.name;
        return { ...s, ref: m.source.id, name: rec.name };
    });

    return { setup: { ...setup, stack: { ...setup.stack, items, sounds } }, found, missing };
}

/** The ids a scene shows that are not found, from rematch's `missing`. */
export function sceneMissing(scene, missing) {
    if (!scene || !missing?.length) return [];
    const lost = missing.filter((m) => m.scope === 'item');
    return (scene.layers || []).map((l) => lost.find((m) => m.id === l.id)).filter(Boolean);
}

/* ── the first start of 2.23 ───────────────────────────────────────────── */

const oldSetup = ({ id, name, app, rawStack, rawStudio, screen, names, now }) => {
    const stack = materialiseScenes(rawStack ? normaliseStack(rawStack) : stackFromSettings(rawStudio, { screen }), names);
    return normaliseSetup({ id, name, app, usedAt: now, stack, settings: normaliseStudio(rawStudio) }, { app });
};

/**
 * The first start of 2.23 (A4.2): the Recorder's kept stack and settings
 * become "My recording setup"; Live Broadcast's, when they differ, "My live
 * setup"; with nothing kept, "My setup" with the four quick scenes. Nothing
 * here writes: the old keys stay as they are, so a step back to 2.22 still
 * has them. Answers { setups, index, global }.
 */
export function migrate({ recorderStack = null, recorderStudio = null, broadcastStack = null, broadcastStudio = null, screen = null, hasCamera = true, names = {}, now = 0, label = {} } = {}) {
    const setups = [];
    const say = { recording: label.recording || 'My recording setup', live: label.live || 'My live setup', first: label.first || 'My setup' };
    const hasRecorder = Boolean(recorderStack || recorderStudio);
    const hasBroadcast = Boolean(broadcastStack || broadcastStudio);
    let recId = '';
    let liveId = '';
    if (hasRecorder) {
        recId = 'setup-recording';
        setups.push(oldSetup({ id: recId, name: say.recording, app: 'recorder', rawStack: recorderStack, rawStudio: recorderStudio, screen, names, now }));
        liveId = recId;
    }
    if (hasBroadcast) {
        const live = oldSetup({ id: 'setup-live', name: say.live, app: 'broadcast', rawStack: broadcastStack, rawStudio: broadcastStudio, screen, names, now });
        const same = hasRecorder && JSON.stringify([live.stack, live.settings]) === JSON.stringify([setups[0].stack, setups[0].settings]);
        if (!same) { setups.push(live); liveId = live.id; }
        if (!hasRecorder) recId = live.id;
    }
    if (!setups.length) {
        const first = newSetup({ id: 'setup-first', name: say.first, app: 'recorder', now, screen, hasCamera, names });
        setups.push(first);
        recId = first.id;
        liveId = first.id;
    }
    let index = normaliseIndex({});
    for (const s of setups) index = withSetup(index, s);
    index = { ...index, current: { recorder: recId, broadcast: liveId } };
    return { setups, index, global: globalSettings(recorderStudio || broadcastStudio || {}) };
}

/* ── the file ──────────────────────────────────────────────────────────── */

const SAFE_ENTRY = /^media\/[^/\\]+$/;
const safeBase = (p) => (String(p).split(/[\\/]/).pop() || 'file').replace(/[^\p{L}\p{N}._ -]/gu, '_').slice(0, 100) || 'file';

/** Every picture and music file a setup uses: pictures on the frame, Back soon's picture, its music. */
export function mediaPaths(setup) {
    const out = [];
    const add = (p) => { if (p && !out.includes(p)) out.push(p); };
    for (const it of setup.stack.items) if (it.kind === 'picture') add(it.path);
    add(setup.stack.placeholder);
    add(setup.stack.music);
    return out;
}

/** A setup with each media path changed by `fn` (a path or '', from a path). */
function mapMedia(setup, fn) {
    const stack = setup.stack;
    return {
        ...setup,
        stack: {
            ...stack,
            items: stack.items.map((it) => (it.kind === 'picture' && it.path ? { ...it, path: fn(it.path) } : it)),
            placeholder: stack.placeholder ? fn(stack.placeholder) : '',
            music: stack.music ? fn(stack.music) : '',
        },
    };
}

const utf8 = (s) => new TextEncoder().encode(s);

/**
 * A setup as a `.rstudio` file: mimetype first and stored (as .rvid), then
 * setup.json, then its pictures and music under media/. `read` gives a file's
 * bytes, or null when it is gone. `streams` is the streams it goes with (name,
 * kind, address), and no key is ever written, whatever `streams` holds.
 * Answers { bytes, absent }, the files that could not be read and so are not
 * in it.
 */
export async function writeSetupFile(setup, { read, streams = [] }) {
    const entries = [];
    const paths = mediaPaths(setup);
    const absent = [];
    const names = new Map();
    let n = 0;
    for (const p of paths) {
        let bytes = null;
        try { bytes = await read(p); } catch { bytes = null; }
        if (!bytes) { absent.push(p); continue; }
        n += 1;
        const entry = `media/${n}-${safeBase(p)}`;
        names.set(p, entry);
        entries.push({ name: entry, bytes });
    }
    const mapped = normaliseSetup({ ...mapMedia(setup, (p) => names.get(p) || ''), tools: false, streams }, { app: setup.app });
    const json = { ...mapped, streams: mapped.streams };
    return {
        bytes: zipStored([
            { name: 'mimetype', bytes: utf8(SETUP_MIME) },
            { name: 'setup.json', bytes: utf8(JSON.stringify(json, null, 2)) },
            ...entries,
        ]),
        absent,
    };
}

/**
 * A `.rstudio` file read. Answers { setup, media: [{ entry, name, bytes }] }:
 * the setup with its media named by entry, and those entries' bytes, each with
 * a plain file name of its own to be kept under. Only what is in the file's
 * media/ can be named: a path to anywhere else in a file from someone else
 * comes out empty. Throws when the file is not a setup.
 */
export async function readSetupFile(bytes) {
    let parts;
    try { parts = await unzip(bytes); } catch { throw new Error('This is not a Rutba Studio setup: it is not a zip.'); }
    const mime = parts.mimetype ? new TextDecoder().decode(parts.mimetype).trim() : '';
    if (mime !== SETUP_MIME || !parts['setup.json']) throw new Error('This is not a Rutba Studio setup.');
    let raw;
    try { raw = JSON.parse(new TextDecoder().decode(parts['setup.json'])); } catch { throw new Error('This setup is damaged: its contents cannot be read.'); }
    const entries = Object.keys(parts).filter((k) => SAFE_ENTRY.test(k));
    const has = new Set(entries);
    const clean = mapMedia(normaliseSetup(raw), (p) => (has.has(p) ? p : ''));
    const media = entries.map((entry, i) => ({ entry, name: `${i + 1}-${safeBase(entry)}`, bytes: parts[entry] }));
    return { setup: clean, media };
}

/** The setup with each media entry replaced by where it now is: `placed` is { entry: path }. */
export function placeMedia(setup, placed) {
    return mapMedia(setup, (p) => placed[p] || '');
}

/** Whether a scene key is the Back soon scene. */
export const isBackSoon = (key) => key === BACK_SOON;
