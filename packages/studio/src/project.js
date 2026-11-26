/**
 * A video project: the studio document as an editor's timeline (decision D1
 * in docs/STUDIO.md).
 *
 * There is no second model. A project IS a studio document (doc.js) whose
 * patches are the clips: a video clip is a `video` layer with a `timing`
 * window on the project's clock and an `offset` into its source, and the
 * clip's sound is a `sound` layer on the same source, in the same `group`,
 * so picture and sound move, trim and split together until they are
 * unlinked. A still is an `image` layer with `fit`; a title is a `text`
 * layer (or a shape and a text in one group). Every layer carries a `lane`:
 * the track the timeline draws it on, counted from 0 within its section, the
 * pictures (`v`) and the sound (`a`). A picture lane's order is its paint
 * order: `planPatches` hands the renderer the patches lane by lane.
 *
 * Everything here is arithmetic on the document: pure, no DOM, no media, and
 * every function answers a NEW document, so undo is a snapshot and autosave
 * a write. The timeline's numbers (seconds) are the project's clock.
 */

import { emptyDoc, normaliseDoc, nextId } from './doc.js';

/** Below this, two times are the same instant (a frame at 120 fps is 0.008 s). */
export const EPS = 0.004;

/** The shortest a clip may be trimmed to. */
export const MIN_CLIP = 0.1;

/** How long a still stays on the timeline unless the setting says otherwise. */
export const STILL_SECONDS = 5;

const round = (t) => Math.round(t * 1000) / 1000;

// ── the document ─────────────────────────────────────────────────────────────

/** A blank project: landscape, dark, nothing of the slideshow's own. */
export function emptyProject() {
    const doc = emptyDoc('video');
    return {
        ...doc,
        size: 'landscape',
        options: { ...doc.options, aspect: 'landscape', theme: 'dark', fit: 'cover', bare: true, fps: 30, quality: 'high' },
        // What each source is, keyed by the url the clips name: its name,
        // size, length, shape and whether it has sound. A .rvid keeps these as
        // the project's source references.
        sources: {},
        markers: [],
        // Per lane: hidden, locked, muted, and a lane's volume.
        lanes: { v: {}, a: {} },
        settings: { ripple: true, snap: true, stillSeconds: STILL_SECONDS },
        // The In and Out points, when set.
        range: null,
    };
}

/** Anything handed to us, as a project: a stored one, or a blank one. */
export function normaliseProject(raw) {
    const base = emptyProject();
    if (!raw || typeof raw !== 'object') return base;
    const doc = normaliseDoc(raw, 'video');
    return {
        ...base,
        ...doc,
        kind: 'video',
        options: { ...base.options, ...(doc.options || {}), bare: true },
        sources: raw.sources && typeof raw.sources === 'object' ? raw.sources : {},
        markers: Array.isArray(raw.markers) ? raw.markers.filter((m) => m && Number.isFinite(Number(m.t))) : [],
        lanes: { v: { ...(raw.lanes?.v || {}) }, a: { ...(raw.lanes?.a || {}) } },
        settings: { ...base.settings, ...(raw.settings || {}) },
        range: raw.range && Number.isFinite(raw.range.in) && Number.isFinite(raw.range.out) ? raw.range : null,
    };
}

// ── reading the timeline ─────────────────────────────────────────────────────

/** Which half of the timeline a layer sits in: pictures, or sound. */
export const sectionOf = (p) => (p?.type === 'sound' ? 'a' : 'v');

/** A clip's lane, 0 when it has none. */
export const laneOf = (p) => Math.max(0, Math.round(Number(p?.lane) || 0));

/** Every clip: the patches that have a window on the clock. */
export const clipsOf = (doc) => (doc?.patches || []).filter((p) => p && p.timing);

export const clipById = (doc, id) => (doc?.patches || []).find((p) => p.id === id) || null;

export const lengthOf = (p) => (p?.timing ? p.timing.end - p.timing.start : 0);

/**
 * How long the project runs: to the end of its last clip. A layer over the
 * whole film (`whole`) follows the film rather than setting its length, so it
 * counts only when nothing else is on the timeline.
 */
export function projectDuration(doc) {
    let end = 0;
    let over = 0;
    for (const p of clipsOf(doc)) {
        if (p.whole) over = Math.max(over, p.timing.end);
        else end = Math.max(end, p.timing.end);
    }
    return round(end || over);
}

/** The project with each layer over the whole film stretched to it, as a version before 2.28.0 reads it. */
export function stretchWhole(doc) {
    if (!(doc?.patches || []).some((p) => p?.whole && p.timing)) return doc;
    const end = projectDuration(doc);
    if (end <= 0) return doc;
    return { ...doc, patches: doc.patches.map((p) => (p?.whole && p.timing ? { ...p, timing: { start: 0, end } } : p)) };
}

/** The clips on one lane, in time order. */
export function laneClips(doc, section, lane) {
    return clipsOf(doc)
        .filter((p) => sectionOf(p) === section && laneOf(p) === lane)
        .sort((a, b) => a.timing.start - b.timing.start || a.timing.end - b.timing.end);
}

/** How many lanes each section has in use (at least one picture lane and one sound lane). */
export function laneCount(doc) {
    let v = 0;
    let a = 0;
    for (const p of clipsOf(doc)) {
        if (sectionOf(p) === 'a') a = Math.max(a, laneOf(p) + 1);
        else v = Math.max(v, laneOf(p) + 1);
    }
    return { v: Math.max(1, v), a: Math.max(1, a) };
}

// ── tracks: insert, move, split (owner, 2026-10-16) ─────────────────────────

/** A section's lane states with each lane number put through `map` (null drops it). */
function remapLaneStates(doc, section, map) {
    const was = doc.lanes?.[section] || {};
    const next = {};
    for (const [k, st] of Object.entries(was)) {
        const to = map(Number(k));
        if (to !== null && to !== undefined && st) next[to] = st;
    }
    return { ...(doc.lanes || { v: {}, a: {} }), [section]: next };
}

/** Every clip of a section with its lane put through `map`, and the lanes' states with them. */
function remapLanes(doc, section, map) {
    const patches = doc.patches.map((p) => (p && p.timing && sectionOf(p) === section ? { ...p, lane: map(laneOf(p)) } : p));
    return { ...doc, patches, lanes: remapLaneStates(doc, section, map) };
}

/** A new empty track at `lane` of a section: the tracks from there up move one up. */
export function insertLane(doc, section, lane) {
    const at = Math.max(0, Math.round(lane) || 0);
    return remapLanes(doc, section, (n) => (n >= at ? n + 1 : n));
}

/**
 * A track moved to `to` with its clips and its state (hidden, locked, its
 * volume): the tracks between close up behind it. For pictures, a higher
 * track is drawn over a lower one.
 */
export function moveLane(doc, section, from, to) {
    const count = laneCount(doc)[section];
    const a = Math.max(0, Math.round(from));
    const b = Math.max(0, Math.min(count - 1, Math.round(to)));
    if (a === b || a >= count) return doc;
    return remapLanes(doc, section, (n) => {
        if (n === a) return b;
        if (a < b && n > a && n <= b) return n - 1;
        if (b < a && n >= b && n < a) return n + 1;
        return n;
    });
}

/**
 * A track split in two at `t` (owner: "split a layer into two so multiple
 * clips can be brought next to each other"): a new track just above it, and
 * the clips from `t` on moved up onto it, a clip under `t` split there
 * first. Their sound or picture stays where it is, in step.
 */
export function splitLane(doc, section, lane, t) {
    const under = laneClips(doc, section, lane).filter((p) => t > p.timing.start + EPS && t < p.timing.end - EPS);
    let out = under.length ? splitClips(doc, under.map((p) => p.id), t).doc : doc;
    const later = laneClips(out, section, lane).filter((p) => p.timing.start >= t - EPS).map((p) => p.id);
    if (!later.length) return doc;
    out = insertLane(out, section, lane + 1);
    const moving = new Set(later);
    return { ...out, patches: out.patches.map((p) => (p && moving.has(p.id) ? { ...p, lane: lane + 1 } : p)) };
}

/** Whether a clip is on a locked track. */
export function onLockedLane(doc, p) {
    return Boolean(p && p.timing && doc.lanes?.[sectionOf(p)]?.[laneOf(p)]?.locked);
}

/**
 * An edit that changed or removed a clip on a locked track, or put a clip
 * onto one: the clips concerned, empty when the locks were kept. Lane tools
 * that renumber whole tracks are not edits of their clips, so `lanesMoved`
 * compares clips leaving their lane number out.
 */
export function brokenLocks(before, after, { lanesMoved = false } = {}) {
    const strip = (p) => (lanesMoved ? { ...p, lane: undefined } : p);
    const now = new Map((after.patches || []).filter(Boolean).map((p) => [p.id, p]));
    const was = new Map((before.patches || []).filter(Boolean).map((p) => [p.id, p]));
    const broken = [];
    for (const p of was.values()) {
        if (!onLockedLane(before, p)) continue;
        const q = now.get(p.id);
        if (!q || JSON.stringify(strip(q)) !== JSON.stringify(strip(p))) broken.push(p.id);
    }
    for (const q of now.values()) {
        if (!was.has(q.id) && onLockedLane(after, q)) broken.push(q.id);
    }
    return broken;
}

/** The other members of a clip's group (its sound, its picture), not the clip itself. */
export function partnersOf(doc, id) {
    const clip = clipById(doc, id);
    if (!clip?.group) return [];
    return clipsOf(doc).filter((p) => p.group === clip.group && p.id !== id);
}

/** A clip and its group, ids only. */
export function withPartners(doc, ids) {
    const out = new Set();
    for (const id of ids) {
        if (!clipById(doc, id)) continue;
        out.add(id);
        for (const p of partnersOf(doc, id)) out.add(p.id);
    }
    return [...out];
}

/** The clips whose window holds `t`, topmost picture lane first, then sound. */
export function clipsAt(doc, t) {
    return clipsOf(doc)
        .filter((p) => t >= p.timing.start - EPS && t <= p.timing.end + EPS)
        .sort((a, b) => (sectionOf(a) === sectionOf(b) ? laneOf(b) - laneOf(a) : sectionOf(a) === 'v' ? -1 : 1));
}

/** A fresh group id, unused in the document. */
function nextGroup(doc, taken = new Set()) {
    const used = new Set([...(doc.patches || []).map((p) => p.group).filter(Boolean), ...taken]);
    let n = 1;
    while (used.has(`g${n}`)) n += 1;
    return `g${n}`;
}

/** Fresh ids for `count` layers of `type`, none of them in the document. */
function freshIds(doc, type, count, taken = new Set()) {
    const ids = [];
    const scratch = { patches: [...(doc.patches || []), ...[...taken].map((id) => ({ id }))] };
    for (let i = 0; i < count; i++) {
        const id = nextId(scratch, type);
        ids.push(id);
        scratch.patches.push({ id });
    }
    return ids;
}

// ── writing ──────────────────────────────────────────────────────────────────

/** Replace some patches by id (a map of id → new patch or null to remove). */
function rewrite(doc, changes) {
    const patches = [];
    for (const p of doc.patches) {
        if (!changes.has(p.id)) { patches.push(p); continue; }
        const next = changes.get(p.id);
        if (next) patches.push(next);
    }
    return { ...doc, patches };
}

// A length given by hand is kept: a bar that covered the whole film no longer follows it.
const setTiming = (p, start, end) => {
    const { whole, ...rest } = p;
    return { ...rest, timing: { start: round(start), end: round(end) } };
};

/**
 * A clip whose start moved `dt` seconds into its own content (a split's
 * second half, a trimmed start): its keys and its automatic zooms are in its
 * own time, so they move back by as much, and what lies before its new start
 * is dropped, a key held at the edge so the motion carries on where it was.
 */
function shiftLocal(p, dt) {
    if (Math.abs(dt) < EPS || (!p.keys && !p.autoZoom)) return p;
    const out = { ...p };
    if (p.keys) {
        const keys = {};
        for (const [prop, list] of Object.entries(p.keys)) {
            if (!Array.isArray(list) || !list.length) continue;
            const moved = list.map((k) => ({ ...k, t: round(k.t - dt) }));
            const before = moved.filter((k) => k.t < 0);
            const after = moved.filter((k) => k.t >= 0);
            if (before.length && after.length && after[0].t > 0) {
                // The value at the new start, held there, so nothing jumps.
                const a = before[before.length - 1];
                const b = after[0];
                after.unshift({ t: 0, v: round(a.v + (b.v - a.v) * ((0 - a.t) / Math.max(1e-6, b.t - a.t))) });
            } else if (before.length && !after.length) {
                after.push({ t: 0, v: before[before.length - 1].v });
            }
            keys[prop] = after;
        }
        out.keys = keys;
    }
    if (p.autoZoom) {
        out.autoZoom = p.autoZoom
            .map((z) => ({ ...z, start: round(z.start - dt), end: round(z.end - dt), ...(z.points ? { points: z.points.map((q) => [round(q[0] - dt), q[1], q[2]]) } : {}) }))
            .filter((z) => z.end > 0)
            .map((z) => ({ ...z, start: Math.max(0, z.start) }));
        if (!out.autoZoom.length) delete out.autoZoom;
    }
    return out;
}

/** Move clips (and nothing else) by `dt` seconds on the clock. */
function shift(doc, ids, dt) {
    if (!ids.length || Math.abs(dt) < EPS) return doc;
    const changes = new Map();
    for (const id of ids) {
        const p = clipById(doc, id);
        if (p) changes.set(id, setTiming(p, Math.max(0, p.timing.start + dt), Math.max(0, p.timing.end + dt)));
    }
    return rewrite(doc, changes);
}

/**
 * The clips after `t` on a lane, with their partners: what a ripple moves.
 * `except` are left where they are (the clip being trimmed).
 */
function rippleSet(doc, section, lane, t, except = []) {
    const skip = new Set(except);
    const ids = laneClips(doc, section, lane)
        .filter((p) => !skip.has(p.id) && p.timing.start >= t - EPS)
        .map((p) => p.id);
    return withPartners(doc, ids).filter((id) => !skip.has(id));
}

// ── adding media (V1) ────────────────────────────────────────────────────────

/**
 * Put a source on the timeline, at the end of the main lane unless `at` and
 * `lane` say otherwise.
 *
 * `media`: { url, name, kind: 'video' | 'audio' | 'image', duration, hasAudio,
 * width, height, size }. A video becomes its picture on a picture lane and,
 * when it has sound, its sound on the sound lane below, in one group; a still
 * a picture of `settings.stillSeconds`; a sound file a sound clip.
 * Answers { doc, ids }: the new layers' ids.
 */
export function addMedia(doc, media, { at = null, lane = null } = {}) {
    const kind = media.kind === 'audio' ? 'audio' : media.kind === 'image' ? 'image' : 'video';
    const section = kind === 'audio' ? 'a' : 'v';
    const useLane = lane ?? 0;
    const length = kind === 'image'
        ? Math.max(MIN_CLIP, Number(doc.settings?.stillSeconds) || STILL_SECONDS)
        : Math.max(MIN_CLIP, Number(media.duration) || STILL_SECONDS);
    const start = at ?? laneEnd(doc, section, useLane);
    const sources = { ...(doc.sources || {}) };
    sources[media.url] = {
        name: media.name || '',
        kind,
        duration: kind === 'image' ? null : (Number(media.duration) || null),
        hasAudio: kind === 'audio' ? true : Boolean(media.hasAudio),
        width: media.width || null,
        height: media.height || null,
        size: media.size ?? null,
    };
    let next = { ...doc, sources };
    const name = media.name || '';
    const ids = [];
    const timing = { start: round(start), end: round(start + length) };
    if (kind === 'audio') {
        const [id] = freshIds(next, 'sound', 1);
        next = { ...next, patches: [...next.patches, { id, type: 'sound', name, url: media.url, offset: 0, volume: 1, timing, lane: useLane, enter: { kind: 'fade', seconds: 0 }, exit: { kind: 'fade', seconds: 0 } }] };
        ids.push(id);
    } else if (kind === 'image') {
        const [id] = freshIds(next, 'image', 1);
        next = { ...next, patches: [...next.patches, { id, type: 'image', name, url: media.url, fit: 'cover', timing, lane: useLane }] };
        ids.push(id);
    } else {
        const group = media.hasAudio ? nextGroup(next) : undefined;
        const [vid] = freshIds(next, 'video', 1);
        const picture = { id: vid, type: 'video', name, url: media.url, offset: 0, timing, lane: useLane, fit: 'cover', ...(group ? { group } : {}) };
        next = { ...next, patches: [...next.patches, picture] };
        ids.push(vid);
        if (media.hasAudio) {
            const [sid] = freshIds(next, 'sound', 1);
            next = { ...next, patches: [...next.patches, { id: sid, type: 'sound', name, url: media.url, offset: 0, volume: 1, timing, lane: 0, group, enter: { kind: 'fade', seconds: 0 }, exit: { kind: 'fade', seconds: 0 } }] };
            ids.push(sid);
        }
    }
    return { doc: next, ids };
}

/** Where a lane's last clip ends. */
export function laneEnd(doc, section, lane) {
    let end = 0;
    for (const p of laneClips(doc, section, lane)) end = Math.max(end, p.timing.end);
    return round(end);
}

/** How long a file lasts on the timeline once added. */
const mediaLength = (doc, media) => (media.kind === 'image'
    ? Math.max(MIN_CLIP, Number(doc.settings?.stillSeconds) || STILL_SECONDS)
    : Math.max(MIN_CLIP, Number(media.duration) || STILL_SECONDS));

/** Whether nothing on a lane covers `start` to `start + length`, but the clips `exclude` names. */
const laneFree = (doc, section, lane, start, length, exclude = new Set()) => laneClips(doc, section, lane)
    .every((p) => exclude.has(p.id) || start + length <= p.timing.start + EPS || start >= p.timing.end - EPS);

/** The lowest sound lane free from `start` for `length`. */
export function freeSoundLane(doc, start, length, exclude = new Set()) {
    let lane = 0;
    while (!laneFree(doc, 'a', lane, start, length, exclude)) lane++;
    return lane;
}

/**
 * The lane music goes on: the lowest one already holding sound files of
 * their own (not a clip's sound), or else the lowest with no sound at all.
 */
export function musicLane(doc) {
    const sounds = clipsOf(doc).filter((p) => p.type === 'sound');
    const own = sounds.filter((p) => !p.group).map(laneOf);
    if (own.length) return Math.min(...own);
    const used = new Set(sounds.map(laneOf));
    let lane = 0;
    while (used.has(lane)) lane++;
    return lane;
}

/**
 * Make room on a lane: a clip there under `at` is split at it, and every
 * clip on the lane from `at` on, with its sound or picture, moves `by`
 * seconds later.
 */
export function pushLater(doc, section, lane, at, by) {
    if (!(by > 0)) return doc;
    const under = laneClips(doc, section, lane).filter((p) => at > p.timing.start + EPS && at < p.timing.end - EPS);
    let out = under.length ? splitClips(doc, under.map((p) => p.id), at).doc : doc;
    const ids = withPartners(out, laneClips(out, section, lane).filter((p) => p.timing.start >= at - EPS).map((p) => p.id));
    const changes = new Map();
    for (const id of ids) {
        const p = clipById(out, id);
        changes.set(id, setTiming(p, p.timing.start + by, p.timing.end + by));
    }
    return rewrite(out, changes);
}

/**
 * Add Media, with the choice of where (owner, 2026-10-16: "the insert media
 * should have choice to add layer or next to existing layer"). `list` is the
 * files as addMedia takes them, in order; pictures (videos and stills) go
 * one after another, and sound files one after another on a lane of their
 * own:
 * - 'end': after the clips already on picture lane `lane`; sound files after
 *   the music already there (musicLane), never spread over a lane each;
 * - 'layer': on a new track from `at`, above the picture tracks there are,
 *   and sound files on a new sound track below;
 * - 'insert': at `at` on picture lane `lane`, the clips there from `at` on
 *   moved later to make room (a clip under `at` split there first); sound
 *   files from `at` on the first sound lane free for them.
 * A video's own sound lands on the first sound lane free for its window.
 * Answers { doc, ids }.
 */
export function placeMedia(doc, list, { where = 'end', lane = 0, at = 0 } = {}) {
    const pictures = list.filter((m) => m.kind !== 'audio');
    const sounds = list.filter((m) => m.kind === 'audio');
    const total = (items) => items.reduce((s, m) => s + mediaLength(doc, m), 0);
    const counts = laneCount(doc);
    const has = (section) => clipsOf(doc).some((p) => sectionOf(p) === section);
    let out = doc;
    let vLane = Math.max(0, Math.round(lane) || 0);
    let vAt;
    let aLane;
    let aAt;
    const from = round(Math.max(0, Number(at) || 0));
    if (where === 'layer') {
        vLane = has('v') ? counts.v : 0;
        aLane = has('a') ? counts.a : 0;
        vAt = from;
        aAt = from;
    } else if (where === 'insert') {
        out = pushLater(out, 'v', vLane, from, total(pictures));
        vAt = from;
        aLane = freeSoundLane(out, from, total(sounds));
        aAt = from;
    } else {
        vAt = laneEnd(doc, 'v', vLane);
        aLane = musicLane(doc);
        aAt = laneEnd(doc, 'a', aLane);
    }
    const ids = [];
    for (const media of pictures) {
        const r = addMedia(out, media, { at: vAt, lane: vLane });
        out = r.doc;
        // The clip's own sound: on a lane where it covers nothing.
        for (const id of r.ids) {
            const p = clipById(out, id);
            if (p?.type !== 'sound') continue;
            const length = p.timing.end - p.timing.start;
            if (laneFree(out, 'a', laneOf(p), p.timing.start, length, new Set([id]))) continue;
            const free = freeSoundLane(out, p.timing.start, length, new Set([id]));
            out = rewrite(out, new Map([[id, { ...p, lane: free }]]));
        }
        ids.push(...r.ids);
        vAt = round(vAt + mediaLength(doc, media));
    }
    for (const media of sounds) {
        const r = addMedia(out, media, { at: aAt, lane: aLane });
        out = r.doc;
        ids.push(...r.ids);
        aAt = round(aAt + mediaLength(doc, media));
    }
    return { doc: out, ids };
}

/**
 * A file's sound alone onto the timeline (Video's audio library): a sound
 * clip from `at`, on the first sound lane free for it. A video file gives its
 * sound and no picture. What the project already knows of the file (a video
 * already on the timeline as a video) stays as it is.
 * `media`: { url, name, duration, size }. Answers { doc, ids }.
 */
export function addSoundOnly(doc, media, { at = 0 } = {}) {
    const length = Math.max(MIN_CLIP, Number(media.duration) || STILL_SECONDS);
    const start = round(Math.max(0, Number(at) || 0));
    const known = doc.sources?.[media.url];
    const r = addMedia(doc, { ...media, kind: 'audio', hasAudio: true }, { at: start, lane: freeSoundLane(doc, start, length) });
    return known ? { doc: { ...r.doc, sources: { ...r.doc.sources, [media.url]: known } }, ids: r.ids } : r;
}

// ── split (V2) ───────────────────────────────────────────────────────────────

/**
 * Split clips at `t`: each becomes two, the second starting at `t` with its
 * source offset moved on by the same amount, so nothing plays differently.
 * A clip's group is split with it, and the second halves form a new group.
 * Clips whose window does not hold `t` strictly inside are left as they are.
 * Answers { doc, ids: the new second halves }.
 */
export function splitClips(doc, ids, t) {
    const targets = withPartners(doc, ids)
        .map((id) => clipById(doc, id))
        .filter((p) => p && t > p.timing.start + EPS && t < p.timing.end - EPS);
    if (!targets.length) return { doc, ids: [] };
    const groups = new Map();
    const taken = new Set();
    const changes = new Map();
    const inserts = new Map();
    const made = [];
    for (const p of targets) {
        let group;
        if (p.group) {
            if (!groups.has(p.group)) groups.set(p.group, nextGroup(doc, new Set(groups.values())));
            group = groups.get(p.group);
        }
        const [id] = freshIds(doc, p.type || 'layer', 1, taken);
        taken.add(id);
        const elapsed = t - p.timing.start;
        const first = {
            ...setTiming(p, p.timing.start, t),
            exit: p.type === 'sound' ? { kind: 'fade', seconds: 0 } : { kind: 'none', seconds: 0 },
        };
        const second = {
            ...shiftLocal(setTiming(p, t, p.timing.end), elapsed),
            id,
            ...(group ? { group } : {}),
            enter: p.type === 'sound' ? { kind: 'fade', seconds: 0 } : { kind: 'none', seconds: 0 },
            ...(p.offset !== undefined || p.type === 'video' || p.type === 'sound' ? { offset: round((Number(p.offset) || 0) + elapsed) } : {}),
        };
        delete second.transition;
        changes.set(p.id, first);
        inserts.set(p.id, second);
        made.push(id);
    }
    const patches = [];
    for (const p of doc.patches) {
        patches.push(changes.get(p.id) || p);
        if (inserts.has(p.id)) patches.push(inserts.get(p.id));
    }
    return { doc: { ...doc, patches }, ids: made };
}

/**
 * Split (S): the selected clips at `t`, or, with none selected that hold
 * `t`, the clip under the playhead on the main picture lane, or the
 * topmost sound clip when no picture is there.
 */
export function splitAt(doc, t, selected = []) {
    const holding = selected.filter((id) => {
        const p = clipById(doc, id);
        return p && t > p.timing.start + EPS && t < p.timing.end - EPS;
    });
    if (holding.length) return splitClips(doc, holding, t);
    const under = clipsAt(doc, t).filter((p) => t > p.timing.start + EPS && t < p.timing.end - EPS);
    // The main picture lane first: a title above the clip is not what S cuts.
    const pick = under.find((p) => sectionOf(p) === 'v' && laneOf(p) === 0) || under.find((p) => sectionOf(p) === 'v') || under[0];
    return pick ? splitClips(doc, [pick.id], t) : { doc, ids: [] };
}

/** Split All Tracks (Shift+S): every clip that holds `t`, on every lane. */
export function splitAll(doc, t) {
    return splitClips(doc, clipsAt(doc, t).map((p) => p.id), t);
}

// ── trim (V3) and ripple trim (V4) ───────────────────────────────────────────

/** The speeds a clip can be set to: a quarter to four times. */
export const SPEEDS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 4];
/** A clip's speed: 1 unless it was set. */
export function speedOf(p) {
    const s = Number(p?.speed);
    return s > 0 ? s : 1;
}

/**
 * A clip's speed, and its group's (its sound, its pointer): the same stretch
 * of its source, shown faster or slower, so its length on the clock is that
 * stretch over the speed. What follows on its lane moves with its end, as a
 * ripple trim does, so nothing is covered and no gap opens.
 */
export function setSpeed(doc, id, speed) {
    const clip = clipById(doc, id);
    if (!clip || !['video', 'sound', 'pointer'].includes(clip.type)) return doc;
    const next = Math.max(0.25, Math.min(4, Number(speed) || 1));
    if (Math.abs(speedOf(clip) - next) < EPS) return doc;
    const ids = withPartners(doc, [id]);
    const changes = new Map();
    for (const pid of ids) {
        const p = clipById(doc, pid);
        if (!['video', 'sound', 'pointer'].includes(p.type)) continue;
        const length = Math.max(MIN_CLIP, (lengthOf(p) * speedOf(p)) / next);
        changes.set(pid, { ...setTiming(p, p.timing.start, p.timing.start + length), speed: next });
    }
    const delta = changes.get(id).timing.end - clip.timing.end;
    let out = rewrite(doc, changes);
    if (Math.abs(delta) > EPS) out = shift(out, rippleSet(doc, sectionOf(clip), laneOf(clip), clip.timing.start + EPS * 2, ids), delta);
    return out;
}

/**
 * Replace Media: a picture clip or a video clip shown from another file of
 * the same kind (`media`, as addMedia takes it), keeping its place, length,
 * look and transitions. A video's own sound follows it to the new file; a
 * video shorter than the clip leaves the clip as long as the new file is.
 * A starter's stand-in so replaced is a stand-in no more.
 */
export function replaceMedia(doc, id, media) {
    const clip = clipById(doc, id);
    if (!clip || !media?.url || !['video', 'image'].includes(clip.type)) return doc;
    const kind = media.kind === 'image' ? 'image' : media.kind === 'audio' ? null : 'video';
    if (kind !== clip.type) return doc;
    const sources = { ...(doc.sources || {}), [media.url]: {
        name: media.name || '', kind,
        duration: kind === 'image' ? null : (Number(media.duration) || null),
        hasAudio: kind === 'video' && Boolean(media.hasAudio),
        width: media.width || null, height: media.height || null, size: media.size ?? null,
    } };
    const changes = new Map();
    const name = media.name || clip.name;
    if (kind === 'image') {
        const { standIn, ...rest } = clip;
        changes.set(id, { ...rest, url: media.url, name });
        return rewrite({ ...doc, sources }, changes);
    }
    const shown = lengthOf(clip) * speedOf(clip);
    const room = Number(media.duration) > 0 ? Number(media.duration) : Infinity;
    const offset = round(Math.max(0, Math.min(clip.offset || 0, room - shown)));
    const length = Math.min(lengthOf(clip), (room - offset) / speedOf(clip));
    for (const pid of withPartners(doc, [id])) {
        const p = clipById(doc, pid);
        if (p.type !== 'video' && p.type !== 'sound') continue;
        if (p.type === 'sound' && (p.url !== clip.url || !media.hasAudio)) continue;
        const { reversed, ...rest } = p;
        changes.set(pid, { ...setTiming(rest, p.timing.start, p.timing.start + Math.max(MIN_CLIP, length)), url: media.url, offset, name });
    }
    // A sound of the old file's with no sound in the new one goes.
    const gone = media.hasAudio ? [] : withPartners(doc, [id]).filter((pid) => { const p = clipById(doc, pid); return p.type === 'sound' && p.url === clip.url; });
    const out = rewrite({ ...doc, sources }, changes);
    return gone.length ? { ...out, patches: out.patches.filter((p) => !gone.includes(p.id)) } : out;
}

/**
 * Reverse: a video clip and its own sound played backward, from a copy of
 * their stretch of the source written backward (`media`, as addMedia takes
 * it). The clip keeps its place, length and speed; a pointer recorded with it
 * is drawn backward too. Asked again, the clip goes back to its own source
 * showing the same stretch, trimmed as it was trimmed while reversed, with no
 * copy needed. `reversed` keeps where the copy came from: { url, offset, span }.
 */
export function setReversed(doc, id, media = null) {
    const clip = clipById(doc, id);
    if (!clip || clip.type !== 'video') return doc;
    const back = Boolean(clip.reversed);
    if (!back && !media?.url) return doc;
    const span = lengthOf(clip) * speedOf(clip);
    let out = doc;
    if (!back) {
        out = { ...doc, sources: { ...(doc.sources || {}), [media.url]: {
            name: media.name || '',
            kind: 'video',
            duration: Number(media.duration) || null,
            hasAudio: Boolean(media.hasAudio),
            width: media.width || null,
            height: media.height || null,
            size: media.size ?? null,
        } } };
    }
    const changes = new Map();
    for (const pid of withPartners(doc, [id])) {
        const p = clipById(doc, pid);
        if (p.type === 'pointer') {
            // Over the copy while reversed (a pointer finds its screen by its file), over the source after.
            if (p.screenUrl !== clip.url) continue;
            const { backward, ...rest } = p;
            changes.set(pid, back ? { ...rest, screenUrl: clip.reversed.url } : { ...rest, backward: true, screenUrl: media.url });
            continue;
        }
        if (p.type !== 'video' && p.type !== 'sound') continue;
        if (back) {
            if (!p.reversed) continue;
            const { reversed, ...rest } = p;
            const shown = lengthOf(p) * speedOf(p);
            const offset = Math.max(0, (reversed.offset || 0) + (reversed.span || 0) - (p.offset || 0) - shown);
            changes.set(pid, { ...rest, url: reversed.url, offset: round(offset) });
        } else {
            // Its own sound, as long as the picture: a sound trimmed apart, or another file, keeps playing forward.
            if (p.url !== clip.url || Math.abs((p.offset || 0) - (clip.offset || 0)) > EPS || Math.abs(lengthOf(p) - lengthOf(clip)) > EPS) continue;
            changes.set(pid, { ...p, url: media.url, offset: 0, reversed: { url: p.url, offset: p.offset || 0, span: round(span) } });
        }
    }
    return rewrite(out, changes);
}

/** The longest a clip may run from `offset`: its source's length, when known. */
function sourceRoom(doc, p) {
    const d = Number(doc.sources?.[p.url]?.duration);
    return p.type === 'video' || p.type === 'sound' ? (Number.isFinite(d) && d > 0 ? d : Infinity) : Infinity;
}

/**
 * Trim a clip's start to `t` on the clock. The source offset rolls with it,
 * so the picture at every remaining instant is unchanged. Its group trims
 * with it. With `ripple`, the clip keeps its place and what follows on its
 * lane closes up (Premiere's ripple trim); without, a gap opens before it.
 */
export function trimStart(doc, id, t, { ripple = false } = {}) {
    const clip = clipById(doc, id);
    if (!clip) return doc;
    const ids = withPartners(doc, [id]);
    // How far the start may move: not past the end less the shortest clip,
    // and not before the start of the source.
    let delta = t - clip.timing.start;
    for (const pid of ids) {
        const p = clipById(doc, pid);
        delta = Math.min(delta, lengthOf(p) - MIN_CLIP);
        if (p.offset !== undefined || p.type === 'video' || p.type === 'sound') delta = Math.max(delta, -(Number(p.offset) || 0));
        delta = Math.max(delta, -p.timing.start);
    }
    if (Math.abs(delta) < EPS) return doc;
    const changes = new Map();
    for (const pid of ids) {
        const p = clipById(doc, pid);
        const hasOffset = p.offset !== undefined || p.type === 'video' || p.type === 'sound';
        const next = { ...shiftLocal(setTiming(p, p.timing.start + delta, p.timing.end), delta), ...(hasOffset ? { offset: round((Number(p.offset) || 0) + delta) } : {}) };
        changes.set(pid, ripple ? setTiming(next, p.timing.start, p.timing.end - delta) : next);
    }
    let out = rewrite(doc, changes);
    if (ripple) {
        const moved = rippleSet(doc, sectionOf(clip), laneOf(clip), clip.timing.start + EPS * 2, ids);
        out = shift(out, moved, -delta);
    }
    return out;
}

/** Trim a clip's end to `t` on the clock, with its group; `ripple` moves what follows on its lane by as much. */
export function trimEnd(doc, id, t, { ripple = false } = {}) {
    const clip = clipById(doc, id);
    if (!clip) return doc;
    const ids = withPartners(doc, [id]);
    let delta = t - clip.timing.end;
    for (const pid of ids) {
        const p = clipById(doc, pid);
        delta = Math.max(delta, MIN_CLIP - lengthOf(p));
        const room = sourceRoom(doc, p);
        // The source left, on the clock: a faster clip uses it up sooner.
        if (Number.isFinite(room)) delta = Math.min(delta, (room - (Number(p.offset) || 0)) / speedOf(p) - lengthOf(p));
    }
    if (Math.abs(delta) < EPS) return doc;
    const changes = new Map();
    for (const pid of ids) {
        const p = clipById(doc, pid);
        changes.set(pid, setTiming(p, p.timing.start, p.timing.end + delta));
    }
    let out = rewrite(doc, changes);
    if (ripple) {
        const moved = rippleSet(doc, sectionOf(clip), laneOf(clip), clip.timing.start + EPS * 2, ids);
        out = shift(out, moved, delta);
    }
    return out;
}

// ── delete, ripple delete and gaps (V4) ──────────────────────────────────────

/**
 * Remove clips and their groups. With `ripple`, everything after each
 * removed clip on its lane (and those clips' groups) moves up to close the
 * space it left; without, a gap stays, which Close Gap closes.
 */
export function removeClips(doc, ids, { ripple = false } = {}) {
    const all = withPartners(doc, ids);
    if (!all.length) return doc;
    // Each lane's removed spans, latest first, so earlier shifts do not move
    // the spans still to be closed.
    const spans = [];
    if (ripple) {
        for (const id of ids) {
            const p = clipById(doc, id);
            if (p) spans.push({ section: sectionOf(p), lane: laneOf(p), start: p.timing.start, end: p.timing.end });
        }
        spans.sort((a, b) => b.start - a.start);
    }
    const gone = new Set(all);
    let out = { ...doc, patches: doc.patches.filter((p) => !gone.has(p.id)) };
    for (const s of spans) {
        // Only the length nothing else on the lane still covers closes.
        const covered = laneClips(out, s.section, s.lane).some((p) => p.timing.start < s.end - EPS && p.timing.end > s.start + EPS);
        if (covered) continue;
        out = shift(out, rippleSet(out, s.section, s.lane, s.start + EPS * 2), -(s.end - s.start));
    }
    return out;
}

/** The empty stretch of a lane that holds `t`, or null: { start, end }. */
export function gapAt(doc, section, lane, t) {
    const clips = laneClips(doc, section, lane);
    let prevEnd = 0;
    for (const p of clips) {
        if (t >= prevEnd - EPS && t < p.timing.start - EPS && p.timing.start - prevEnd > EPS) return { start: prevEnd, end: p.timing.start };
        prevEnd = Math.max(prevEnd, p.timing.end);
    }
    return null;
}

/** Close Gap: the empty stretch at `t` on a lane closes, what follows moving up. */
export function closeGap(doc, section, lane, t) {
    const gap = gapAt(doc, section, lane, t);
    if (!gap) return doc;
    return shift(doc, rippleSet(doc, section, lane, gap.end), -(gap.end - gap.start));
}

// ── move, copy, paste, duplicate (V5) ────────────────────────────────────────

/**
 * Where a clip of `length` can start on a lane, nearest to `start`, without
 * covering another clip (`exclude` are the clips being moved).
 */
export function freeStart(doc, section, lane, start, length, exclude = []) {
    const skip = new Set(exclude);
    const others = laneClips(doc, section, lane).filter((p) => !skip.has(p.id));
    const fits = (s) => s >= -EPS && others.every((p) => s + length <= p.timing.start + EPS || s >= p.timing.end - EPS);
    if (fits(start)) return round(Math.max(0, start));
    const candidates = [0, ...others.map((p) => p.timing.end), ...others.map((p) => p.timing.start - length)]
        .filter((s) => s >= -EPS && fits(s));
    candidates.sort((a, b) => Math.abs(a - start) - Math.abs(b - start));
    return round(Math.max(0, candidates[0] ?? laneEnd({ patches: others }, section, lane)));
}

/**
 * Move a clip (and its group) so it starts at `start`, on `lane` when given.
 * Group partners keep their own lanes and move by the same time. The clip
 * lands on the nearest place on its lane where it covers nothing.
 */
export function moveClip(doc, id, { start, lane = null } = {}) {
    const clip = clipById(doc, id);
    if (!clip) return doc;
    const ids = withPartners(doc, [id]);
    const section = sectionOf(clip);
    const toLane = lane === null || lane === undefined ? laneOf(clip) : Math.max(0, Math.round(lane));
    const want = Math.max(0, start ?? clip.timing.start);
    const at = freeStart(doc, section, toLane, want, lengthOf(clip), ids);
    let dt = at - clip.timing.start;
    // Partners must not go below zero either.
    for (const pid of ids) dt = Math.max(dt, -clipById(doc, pid).timing.start);
    const changes = new Map();
    for (const pid of ids) {
        const p = clipById(doc, pid);
        const moved = setTiming(p, p.timing.start + dt, p.timing.end + dt);
        if (pid === id) moved.lane = toLane;
        // A moved clip leaves its transition behind: it no longer follows the clip it faded from.
        if (pid === id && p.transition) {
            delete moved.transition;
            moved.enter = p.type === 'sound' ? { kind: 'fade', seconds: 0 } : { kind: 'none', seconds: 0 };
        }
        changes.set(pid, moved);
    }
    return rewrite(doc, changes);
}

/** The clips to copy: a clip and its group, as patches. */
export function copyClips(doc, ids) {
    const all = withPartners(doc, ids);
    const clips = all.map((id) => clipById(doc, id)).filter(Boolean);
    const from = Math.min(...clips.map((p) => p.timing.start));
    return { from: Number.isFinite(from) ? from : 0, clips: clips.map((p) => JSON.parse(JSON.stringify(p))) };
}

/**
 * Paste copied clips so the earliest starts at `at`, each on its own lane,
 * with fresh ids and groups. Answers { doc, ids }.
 */
export function pasteClips(doc, copied, at) {
    if (!copied?.clips?.length) return { doc, ids: [] };
    const taken = new Set();
    const groups = new Map();
    const fresh = [];
    for (const p of copied.clips) {
        const [id] = freshIds(doc, p.type || 'layer', 1, taken);
        taken.add(id);
        let group;
        if (p.group) {
            if (!groups.has(p.group)) groups.set(p.group, nextGroup(doc, new Set(groups.values())));
            group = groups.get(p.group);
        }
        const copy = { ...p, id, ...(group ? { group } : {}) };
        delete copy.transition;
        fresh.push(copy);
    }
    // Shift the whole set so its earliest clip lands at `at`, then on to the
    // first place where none of it covers a clip already there.
    let dt = at - copied.from;
    for (let tries = 0; tries < 200; tries++) {
        const clash = fresh.find((p) => laneClips(doc, sectionOf(p), laneOf(p)).some((q) => p.timing.start + dt < q.timing.end - EPS && p.timing.end + dt > q.timing.start + EPS));
        if (!clash) break;
        const blocker = laneClips(doc, sectionOf(clash), laneOf(clash)).find((q) => clash.timing.start + dt < q.timing.end - EPS && clash.timing.end + dt > q.timing.start + EPS);
        dt = blocker.timing.end - clash.timing.start;
    }
    const placed = fresh.map((p) => setTiming(p, p.timing.start + dt, p.timing.end + dt));
    return { doc: { ...doc, patches: [...doc.patches, ...placed] }, ids: placed.map((p) => p.id) };
}

/** Duplicate: a copy of the clips straight after the last of them. */
export function duplicateClips(doc, ids) {
    const copied = copyClips(doc, ids);
    if (!copied.clips.length) return { doc, ids: [] };
    const end = Math.max(...copied.clips.map((p) => p.timing.end));
    return pasteClips(doc, copied, end);
}

// ── snapping (V5) ────────────────────────────────────────────────────────────

/**
 * Snap a time to the nearest edge within `threshold` seconds: clip edges
 * (except the clips in `exclude`), the playhead, markers, the In and Out
 * points, zero, and the beats of a music bed when the document has them.
 * Answers { t, to }: `to` the time snapped to, or null.
 */
export function snapTime(doc, t, { threshold = 0.1, exclude = [], playhead = null, beats = null } = {}) {
    const skip = new Set(exclude);
    const points = [0];
    for (const p of clipsOf(doc)) {
        if (skip.has(p.id)) continue;
        points.push(p.timing.start, p.timing.end);
    }
    if (Number.isFinite(playhead)) points.push(playhead);
    for (const m of doc.markers || []) points.push(Number(m.t));
    if (doc.range) points.push(doc.range.in, doc.range.out);
    for (const b of Array.isArray(beats) ? beats : []) points.push(Number(b));
    let best = null;
    for (const q of points) {
        if (!Number.isFinite(q)) continue;
        const d = Math.abs(q - t);
        if (d <= threshold && (best === null || d < Math.abs(best - t))) best = q;
    }
    return best === null ? { t, to: null } : { t: round(best), to: round(best) };
}

// ── groups (V6) ──────────────────────────────────────────────────────────────

/** Detach Audio, or Unlink: the clips of a group go their own ways. */
export function unlinkClips(doc, ids) {
    const all = withPartners(doc, ids);
    const changes = new Map();
    for (const id of all) {
        const p = clipById(doc, id);
        if (p?.group) {
            const { group, ...rest } = p;
            changes.set(id, rest);
        }
    }
    return changes.size ? rewrite(doc, changes) : doc;
}

/** Link: the clips move together from now on (a picture and its sound, relinked). */
export function linkClips(doc, ids) {
    const clips = ids.map((id) => clipById(doc, id)).filter(Boolean);
    if (clips.length < 2) return doc;
    const group = nextGroup(doc);
    const changes = new Map(clips.map((p) => [p.id, { ...p, group }]));
    // Anything that was grouped with one of them keeps its own old group.
    return rewrite(doc, changes);
}

// ── markers (V7) ─────────────────────────────────────────────────────────────

export function addMarker(doc, t, name = '') {
    const markers = [...(doc.markers || [])];
    let n = 1;
    const used = new Set(markers.map((m) => m.id));
    while (used.has(`m${n}`)) n += 1;
    markers.push({ id: `m${n}`, t: round(Math.max(0, t)), name: String(name || '') });
    markers.sort((a, b) => a.t - b.t);
    return { ...doc, markers };
}

export function updateMarker(doc, id, change) {
    return { ...doc, markers: (doc.markers || []).map((m) => (m.id === id ? { ...m, ...change, t: round(Number(change.t ?? m.t)) } : m)).sort((a, b) => a.t - b.t) };
}

export function removeMarker(doc, id) {
    return { ...doc, markers: (doc.markers || []).filter((m) => m.id !== id) };
}

/** The first marker after `t`, or null. */
/** A time as a chapters list writes it: m:ss, or h:mm:ss from an hour. */
function chapterTime(t) {
    const all = Math.max(0, Math.floor(t + 1e-6));
    const h = Math.floor(all / 3600);
    const m = Math.floor((all % 3600) / 60);
    const s = String(all % 60).padStart(2, '0');
    return h ? `${h}:${String(m).padStart(2, '0')}:${s}` : `${m}:${s}`;
}

/**
 * The markers as a chapters list, one "0:00 Title" line each, as video sites
 * read it (S2). The time is that of what is exported: from `from` to `to`
 * when only part is. The list starts at 0:00, with `first` when no marker
 * is there; a marker with no name is called `unnamed(n)`, n its place in
 * the list. No markers, no list.
 */
export function chaptersText(doc, { from = 0, to = Infinity, first = 'Start', unnamed = (n) => `Chapter ${n}` } = {}) {
    const inside = (doc?.markers || []).filter((m) => m.t >= from - EPS && m.t < to - EPS).map((m) => ({ t: m.t - from, name: String(m.name || '').trim() }));
    if (!inside.length) return '';
    inside.sort((a, b) => a.t - b.t);
    if (Math.floor(inside[0].t + 1e-6) > 0) inside.unshift({ t: 0, name: first });
    return inside.map((m, i) => `${chapterTime(m.t)} ${m.name || unnamed(i + 1)}\n`).join('');
}

export function nextMarker(doc, t) {
    return (doc.markers || []).find((m) => m.t > t + EPS) || null;
}

/** The last marker before `t`, or null. */
export function previousMarker(doc, t) {
    const before = (doc.markers || []).filter((m) => m.t < t - EPS);
    return before[before.length - 1] || null;
}

// ── transitions (V15) ────────────────────────────────────────────────────────

/**
 * The transitions, each an incoming clip's enter (and for a push, a dip and a
 * blur the outgoing one's exit). A wipe, a slide and a push are named for the
 * way they travel: plain is leftwards across the frame for a slide and a
 * push, rightwards for a wipe, as they always were; the others say their way.
 */
export const TRANSITIONS = [
    'fade', 'dip', 'wipe', 'slide', 'push', 'zoom',
    'wipe-left', 'wipe-down', 'wipe-up',
    'slide-right', 'slide-up', 'slide-down',
    'push-right', 'push-up', 'push-down',
    'iris', 'blur',
];
const ENTER_KIND = {
    fade: 'fade', dip: 'fade', wipe: 'wipe', slide: 'slide-left', push: 'push', zoom: 'zoom',
    'wipe-left': 'wipe-left', 'wipe-down': 'wipe-down', 'wipe-up': 'wipe-up',
    'slide-right': 'slide-right', 'slide-up': 'slide-up', 'slide-down': 'slide-down',
    'push-right': 'slide-right', 'push-up': 'slide-up', 'push-down': 'slide-down',
    iris: 'iris', blur: 'blur-through',
};
/** What the outgoing clip does while the incoming one arrives, for the transitions that move it. */
const EXIT_KIND = { push: 'push', 'push-right': 'slide-right', 'push-up': 'slide-up', 'push-down': 'slide-down', blur: 'blur-through' };
/** Every transition but the dip overlaps the two clips. */
const OVERLAPS = new Set(TRANSITIONS.filter((k) => k !== 'dip'));

/** The clip on the same lane that a clip follows: the one that ends last at or before it starts (or that it overlaps). */
export function previousOnLane(doc, id) {
    const clip = clipById(doc, id);
    if (!clip) return null;
    const overlap = clip.transition && OVERLAPS.has(clip.transition.kind) ? clip.transition.seconds : 0;
    const before = laneClips(doc, sectionOf(clip), laneOf(clip))
        .filter((p) => p.id !== id && p.timing.start < clip.timing.start - EPS && p.timing.end <= clip.timing.start + overlap + EPS);
    return before.length ? before.reduce((a, b) => (b.timing.end > a.timing.end ? b : a)) : null;
}

/** The sound clip in a picture clip's group, if any. */
const soundOf = (doc, p) => (p?.group ? clipsOf(doc).find((q) => q.group === p.group && q.type === 'sound') || null : null);

/** The clip that follows a clip on its lane with a transition from it, if any: that transition is this clip's way out. */
export function transitionAfter(doc, id) {
    const clip = clipById(doc, id);
    if (!clip) return null;
    return laneClips(doc, sectionOf(clip), laneOf(clip)).find((p) => p.id !== id && p.transition && previousOnLane(doc, p.id)?.id === id) || null;
}

/**
 * Put a transition between a clip and the one before it on its lane, or take
 * it off (`kind: 'none'`). Every transition but the dip overlaps the two
 * clips by the transition's length: the incoming clip (and
 * everything after it on its lane, with their groups) moves up by that much,
 * and the sound crossfades with the picture. A dip goes to black and back
 * without an overlap. The length is held to half of each clip.
 *
 * Between two sound clips of their own (music, a voice; not a picture's
 * sound, which crossfades with its picture) the one transition is the
 * crossfade: they overlap, the first fading out as the second fades in.
 */
export function setTransition(doc, id, { kind = 'fade', seconds = 1 } = {}) {
    let out = clearTransition(doc, id);
    const clip = clipById(out, id);
    if (!clip || kind === 'none' || !TRANSITIONS.includes(kind)) return out;
    const before = previousOnLane(out, id);
    if (!before) return out;
    const gap = clip.timing.start - before.timing.end;
    // Only clips that meet, or nearly: a transition across a gap would fade to nothing.
    if (gap > 0.05) return out;
    const d = round(Math.max(0.1, Math.min(Number(seconds) || 1, lengthOf(clip) / 2, lengthOf(before) / 2)));
    if (clip.type === 'sound' || before.type === 'sound') {
        if (kind !== 'fade' || clip.type !== 'sound' || before.type !== 'sound' || clip.group || before.group) return out;
        // What follows on the lane moves up with it; a picture's sound there stays with its picture.
        const moved = rippleSet(out, 'a', laneOf(clip), clip.timing.start).filter((pid) => !clipById(out, pid).group);
        out = shift(out, moved, -(d + Math.max(0, gap)));
        const incoming = clipById(out, id);
        const outgoing = clipById(out, before.id);
        return rewrite(out, new Map([
            [incoming.id, { ...incoming, enter: { kind: 'fade', seconds: d }, transition: { kind: 'fade', seconds: d } }],
            [outgoing.id, { ...outgoing, exit: { kind: 'fade', seconds: d } }],
        ]));
    }
    const overlaps = OVERLAPS.has(kind);
    if (overlaps) {
        const moved = rippleSet(out, sectionOf(clip), laneOf(clip), clip.timing.start);
        out = shift(out, withPartners(out, moved), -(d + Math.max(0, gap)));
    }
    const incoming = clipById(out, id);
    const outgoing = clipById(out, before.id);
    const ramp = overlaps ? d : round(d / 2);
    const changes = new Map();
    changes.set(incoming.id, { ...incoming, enter: { kind: ENTER_KIND[kind], seconds: ramp }, transition: { kind, seconds: d } });
    const outExit = EXIT_KIND[kind] ? { kind: EXIT_KIND[kind], seconds: d } : kind === 'dip' ? { kind: 'fade', seconds: ramp } : { kind: 'none', seconds: 0 };
    changes.set(outgoing.id, { ...outgoing, exit: outExit });
    // The sound crossfades (or dips) with the picture.
    const inSound = soundOf(out, incoming);
    const outSound = soundOf(out, outgoing);
    if (inSound) changes.set(inSound.id, { ...inSound, enter: { kind: 'fade', seconds: ramp } });
    if (outSound) changes.set(outSound.id, { ...outSound, exit: { kind: 'fade', seconds: ramp } });
    return rewrite(out, changes);
}

/** Take a clip's transition off: the overlap opens again and the ramps go. */
export function clearTransition(doc, id) {
    const clip = clipById(doc, id);
    if (!clip?.transition) return doc;
    const { kind, seconds } = clip.transition;
    const before = previousOnLane(doc, id);
    let out = doc;
    const changes = new Map();
    const { transition, ...rest } = clip;
    changes.set(id, { ...rest, enter: clip.type === 'sound' ? { kind: 'fade', seconds: 0 } : { kind: 'none', seconds: 0 } });
    if (before) changes.set(before.id, { ...before, exit: before.type === 'sound' ? { kind: 'fade', seconds: 0 } : { kind: 'none', seconds: 0 } });
    const inSound = soundOf(doc, clip);
    const outSound = soundOf(doc, before);
    if (inSound) changes.set(inSound.id, { ...inSound, enter: { kind: 'fade', seconds: 0 } });
    if (outSound) changes.set(outSound.id, { ...outSound, exit: { kind: 'fade', seconds: 0 } });
    out = rewrite(out, changes);
    if (clip.type === 'sound') {
        out = shift(out, rippleSet(out, 'a', laneOf(clip), clip.timing.start).filter((pid) => !clipById(out, pid).group), seconds);
    } else if (OVERLAPS.has(kind)) {
        const moved = rippleSet(out, sectionOf(clip), laneOf(clip), clip.timing.start);
        out = shift(out, withPartners(out, moved), seconds);
    }
    return out;
}

// ── titles (V21, V22) ────────────────────────────────────────────────────────

/**
 * The animations a layer comes in and goes out with (D-2026-11-22-14), for
 * every kind of layer: each an envelope kind the renderer applies around the
 * layer's painter, so a clip, a picture, a shape and a chart take them as a
 * title does. Rise and drop travel a little way; left, right, up and down
 * cross the frame, named for the way the layer travels (in from the far side,
 * out through the near one). Type is a title's alone, and only on the way in.
 */
export const ANIMATIONS = {
    none: { kind: 'none', seconds: 0 },
    fade: { kind: 'fade', seconds: 0.5 },
    rise: { kind: 'slide-up', seconds: 0.6, dist: 60 },
    drop: { kind: 'slide-down', seconds: 0.6, dist: 60 },
    type: { kind: 'type-on', seconds: 0 },
    pop: { kind: 'zoom', seconds: 0.45 },
    grow: { kind: 'grow', seconds: 0.5 },
    spin: { kind: 'spin', seconds: 0.6 },
    left: { kind: 'slide-left', seconds: 0.6 },
    right: { kind: 'slide-right', seconds: 0.6 },
    up: { kind: 'slide-up', seconds: 0.6 },
    down: { kind: 'slide-down', seconds: 0.6 },
    wipe: { kind: 'wipe', seconds: 0.6 },
    iris: { kind: 'iris', seconds: 0.6 },
    blur: { kind: 'blur-through', seconds: 0.6 },
};

/** The same list under the name it had when titles alone were animated. */
export const TEXT_ANIMATIONS = ANIMATIONS;

const ANIMATION_OF_KIND = { 'type-on': 'type', zoom: 'pop', grow: 'grow', spin: 'spin', 'slide-left': 'left', push: 'left', 'slide-right': 'right', iris: 'iris', 'blur-through': 'blur', fade: 'fade' };

/** Which animation an envelope is, by name. */
export function animationName(env) {
    if (!env || env.kind === 'none' || !env.kind) return 'none';
    if (env.kind === 'slide-up') return env.dist ? 'rise' : 'up';
    if (env.kind === 'slide-down') return env.dist ? 'drop' : 'down';
    if (String(env.kind).startsWith('wipe')) return 'wipe';
    return ANIMATION_OF_KIND[env.kind] || 'fade';
}

/** The kinds of layer that are not animated in or out: a sound fades by its own lengths, and a fade at the two ends is one already. */
const NOT_ANIMATED = new Set(['sound', 'pointer', 'edges']);

/** The animations a clip can come in with (`edge` 'in') or go out with ('out'), by name, None first. */
export function animationsFor(clip, edge = 'in') {
    if (!clip?.timing || NOT_ANIMATED.has(clip.type)) return [];
    return Object.keys(ANIMATIONS).filter((name) => name !== 'type' || (edge === 'in' && clip.type === 'text'));
}

/**
 * The envelope of an animation by name, at its own length or at `seconds`
 * (held to half the clip when `clip` is given). Typing on has no length of
 * its own: it runs at the pace of its letters.
 */
export function animationEnvelope(name, seconds = null, clip = null) {
    const base = ANIMATIONS[name] || ANIMATIONS.none;
    if (base.kind === 'none' || base.kind === 'type-on') return { ...base };
    const most = clip ? Math.max(0.1, lengthOf(clip) / 2) : Infinity;
    const wanted = seconds === null || seconds === undefined || !Number.isFinite(Number(seconds)) ? base.seconds : Number(seconds);
    return { ...base, seconds: round(Math.max(0.1, Math.min(wanted, most))) };
}

/** The longest an animation lasts on a clip, in seconds: half of it, so its way in and its way out never meet. */
export const animationMost = (clip) => round(Math.max(0.1, lengthOf(clip) / 2));

/**
 * The envelope for a length typed at one end of a clip (`edge` 'in' or
 * 'out'): held to half the clip, as a length chosen from the list is. No
 * length is no animation, and a length where there was no kind is a fade.
 */
export function animationLasting(clip, edge, seconds) {
    const name = animationName(clip?.[edge === 'in' ? 'enter' : 'exit']);
    const wanted = Number(seconds);
    if (!(wanted > 0)) return { kind: 'none', seconds: 0 };
    return animationEnvelope(name === 'none' ? 'fade' : name, wanted, clip);
}

// ── movement over time (D-2026-11-22-14) ─────────────────────────────────────
//
// A clip moves by points set on the timeline. A point is a time inside the
// clip and what the clip looks like there: where it stands, how big it is,
// its angle and how solid it is. Between two points the renderer carries it
// from one to the other. The points are the renderer's own `keys` on the
// layer: { prop: [{ t, v, ease? }] }, with t counted from the clip's start,
// so a clip moved along the timeline carries its movement, and a split or a
// trimmed start keeps it (shiftLocal). Every point holds every property the
// clip moves by, so a point is one thing to add, drag and remove.

/** Two times closer than this are the same point. */
export const MOTION_SNAP = 0.04;

/** How a move arrives at a point: evenly, slow at first, slow at the end, or slow at both. */
export const MOTION_EASES = ['', 'in', 'out', 'in-out'];

const MOTION_DEFAULT = { fx: 0.5, fy: 0.5, rot: 0, opacity: 1, zoom: 1, panX: 0.5, panY: 0.5 };
/** The kinds that lie over the whole frame: they have no place or size to move, only how solid they are. */
const OVER_FRAME = new Set(['adjust', 'gradient', 'progress', 'edges', 'outro']);
const fillsFrame = (p) => (p.type === 'video' && !p.fw) || (p.type === 'image' && Boolean(p.fit));

/**
 * What a point of movement holds for a clip. A layer that stands in a box of
 * its own: its place, the sizes it has, its angle and its opacity. A clip or
 * a still that fills the frame: how far it is zoomed in and where (a slow
 * push in, a pan across a picture), and its opacity; the zoom is left out
 * where the zooms on clicks own it. A layer over the whole frame: its
 * opacity. A sound, a caption and a pointer do not move.
 */
export function motionProps(clip) {
    if (!clip?.timing || clip.type === 'sound' || clip.type === 'pointer' || clip.caption) return [];
    if (OVER_FRAME.has(clip.type)) return ['opacity'];
    if (fillsFrame(clip)) return clip.autoZoom?.length ? ['opacity'] : ['zoom', 'panX', 'panY', 'opacity'];
    const out = ['fx', 'fy'];
    // A QR code's size is not moved: it is drawn again at each size, and one that changes cannot be read.
    if (clip.type !== 'qr') for (const prop of ['fw', 'fh', 'sizeFrac']) if (Number.isFinite(clip[prop]) && clip[prop] > 0) out.push(prop);
    if (clip.type !== 'qr') out.push('rot');
    out.push('opacity');
    return out;
}

const keyList = (clip, prop) => (Array.isArray(clip?.keys?.[prop]) ? clip.keys[prop] : []);
const easeP = (x, ease) => (ease === 'in' ? x * x : ease === 'out' ? 1 - (1 - x) * (1 - x) : ease === 'in-out' ? x * x * (3 - 2 * x) : x);

/** The times of a clip's points, from its start, in order. */
export function motionTimes(clip) {
    const times = [];
    for (const prop of motionProps(clip)) for (const k of keyList(clip, prop)) if (Number.isFinite(k?.t)) times.push(k.t);
    times.sort((a, b) => a - b);
    return times.filter((t, i) => i === 0 || t - times[i - 1] > EPS);
}

/** Whether a clip has points of movement. */
export const isMoving = (clip) => motionTimes(clip).length > 0;

/** The point of a clip at `t` (from its start), or null: its time as kept. */
export function motionPointAt(clip, t) {
    let best = null;
    for (const at of motionTimes(clip)) if (Math.abs(at - t) <= MOTION_SNAP && (best === null || Math.abs(at - t) < Math.abs(best - t))) best = at;
    return best;
}

/** What one property of a clip is at `t` from its start: carried between its points as the renderer carries it, else as it stands. */
export function motionValue(clip, prop, t) {
    const list = keyList(clip, prop);
    const still = Number.isFinite(clip?.[prop]) ? clip[prop] : MOTION_DEFAULT[prop];
    if (!list.length) return still;
    if (t <= list[0].t) return list[0].v;
    const last = list[list.length - 1];
    if (t >= last.t) return last.v;
    for (let i = 1; i < list.length; i++) {
        if (t <= list[i].t) {
            const a = list[i - 1];
            const b = list[i];
            return a.v + (b.v - a.v) * easeP((t - a.t) / Math.max(1e-6, b.t - a.t), b.ease);
        }
    }
    return last.v;
}

/** A clip as it looks at `t` from its start: what moves read at that time, the rest as it is. For the numbers a properties pane shows. */
export function clipAtTime(clip, t) {
    if (!isMoving(clip)) return clip;
    const out = { ...clip };
    for (const prop of motionProps(clip)) {
        const v = motionValue(clip, prop, t);
        if (v !== undefined) out[prop] = round4(v);
    }
    return out;
}

const round4 = (v) => Math.round(v * 10000) / 10000;
const localTime = (clip, t) => round(Math.max(0, Math.min(lengthOf(clip), Number(t) || 0)));

/** A clip's keys with every point holding every property it moves by, each list in order. */
function wholePoints(clip, keys) {
    const props = motionProps(clip);
    const probe = { ...clip, keys };
    const times = motionTimes(probe);
    const out = { ...keys };
    for (const prop of props) {
        const list = keyList(probe, prop);
        const next = times.map((t) => {
            const own = list.find((k) => Math.abs(k.t - t) <= EPS);
            if (own) return own;
            const v = motionValue(probe, prop, t);
            return v === undefined ? null : { t, v: round4(v) };
        }).filter(Boolean);
        if (next.length) out[prop] = next;
        else delete out[prop];
    }
    return out;
}

const withKeys = (clip, keys) => {
    const { keys: old, ...rest } = clip;
    return Object.keys(keys).length ? { ...rest, keys } : rest;
};

/** A clip with a point at `t` from its start, holding what it looks like there, changed by `change` where given. */
function pointSet(clip, t, change = {}) {
    const props = motionProps(clip);
    const at = motionPointAt(clip, t) ?? localTime(clip, t);
    const keys = { ...(clip.keys || {}) };
    for (const prop of props) {
        const v = prop in change && Number.isFinite(change[prop]) ? change[prop] : motionValue(clip, prop, at);
        if (v === undefined) continue;
        const list = keyList(clip, prop);
        const old = list.find((k) => Math.abs(k.t - at) <= EPS);
        const key = { t: at, v: round4(v), ...(old?.ease ? { ease: old.ease } : {}) };
        keys[prop] = [...list.filter((k) => Math.abs(k.t - at) > EPS), key].sort((a, b) => a.t - b.t);
    }
    return withKeys(clip, wholePoints(clip, keys));
}

/**
 * A change to a clip made at `t` from its start, as the grips on the preview
 * and the numbers in the properties pane make it. A clip with no points
 * takes it as it is. A clip that moves takes what moves into the point at
 * that time, made there when there was none, so dragging it at another time
 * is how the next point is set; anything else in the change is the clip's
 * own, as before.
 */
export function changeAtTime(clip, t, change) {
    if (!clip || !change) return clip;
    if (!isMoving(clip)) return { ...clip, ...change };
    const moves = new Set(motionProps(clip));
    const still = {};
    const moved = {};
    for (const [prop, v] of Object.entries(change)) {
        if (moves.has(prop) && Number.isFinite(v)) moved[prop] = v;
        else still[prop] = v;
    }
    let next = keptKeys({ ...clip, ...still });
    const now = new Set(motionProps(next));
    for (const prop of Object.keys(moved)) if (!now.has(prop)) { next = { ...next, [prop]: moved[prop] }; delete moved[prop]; }
    if (Object.keys(moved).length) return pointSet(next, t, moved);
    // Its points hold whatever it moves by now, as it stands.
    return next.keys && isMoving(next) ? withKeys(next, wholePoints(next, next.keys)) : next;
}

/**
 * A clip whose kind of movement changed (a still made to fill the frame, a
 * clip made a picture in picture): the points of what it no longer moves by
 * go, so nothing it cannot be seen to have still moves it. The zooms on
 * clicks keep their own.
 */
function keptKeys(clip) {
    if (!clip.keys) return clip;
    const moves = new Set(motionProps(clip));
    const keys = {};
    for (const [prop, list] of Object.entries(clip.keys)) {
        const zoomOwn = clip.autoZoom?.length && (prop === 'zoom' || prop === 'panX' || prop === 'panY');
        if (moves.has(prop) || zoomOwn) keys[prop] = list;
    }
    return Object.keys(keys).length === Object.keys(clip.keys).length ? clip : withKeys(clip, keys);
}

/** `changeAtTime` on a clip of a project, with `time` on the project's clock. */
export function patchAtTime(doc, id, time, change) {
    const clip = clipById(doc, id);
    if (!clip) return doc;
    return rewrite(doc, new Map([[id, changeAtTime(clip, time - clip.timing.start, change)]]));
}

/** Add a point to a clip at `time` on the project's clock, holding what it looks like there. */
export function addMotionPoint(doc, id, time) {
    const clip = clipById(doc, id);
    if (!clip || !motionProps(clip).length) return doc;
    const t = time - clip.timing.start;
    if (t < -EPS || t > lengthOf(clip) + EPS || motionPointAt(clip, localTime(clip, t)) !== null) return doc;
    return rewrite(doc, new Map([[id, pointSet(clip, t)]]));
}

/**
 * Take the point at `t` (from the clip's start) off a clip. When it was the
 * last, the clip stands still as that point had it.
 */
export function removeMotionPoint(doc, id, t) {
    const clip = clipById(doc, id);
    const at = clip ? motionPointAt(clip, t) : null;
    if (at === null) return doc;
    const props = motionProps(clip);
    const last = motionTimes(clip).length === 1;
    const keys = { ...(clip.keys || {}) };
    const still = {};
    for (const prop of props) {
        const list = keyList(clip, prop);
        if (last) { const k = list.find((x) => Math.abs(x.t - at) <= EPS); if (k) still[prop] = k.v; }
        const rest = list.filter((k) => Math.abs(k.t - at) > EPS);
        if (rest.length) keys[prop] = rest;
        else delete keys[prop];
    }
    return rewrite(doc, new Map([[id, withKeys({ ...clip, ...still }, keys)]]));
}

/** Move a clip's point from `from` to `to` (both from the clip's start), held inside the clip. A point it lands on is replaced. */
export function moveMotionPoint(doc, id, from, to) {
    const clip = clipById(doc, id);
    const at = clip ? motionPointAt(clip, from) : null;
    if (at === null) return doc;
    const target = localTime(clip, to);
    if (Math.abs(target - at) <= EPS) return doc;
    const keys = { ...(clip.keys || {}) };
    for (const prop of motionProps(clip)) {
        const list = keyList(clip, prop);
        const own = list.find((k) => Math.abs(k.t - at) <= EPS);
        if (!own) continue;
        keys[prop] = [...list.filter((k) => k !== own && Math.abs(k.t - target) > MOTION_SNAP), { ...own, t: target }].sort((a, b) => a.t - b.t);
    }
    return rewrite(doc, new Map([[id, withKeys(clip, keys)]]));
}

/** How the move into a clip's point at `t` goes: '' evenly, 'in', 'out' or 'in-out'. */
export function motionEase(clip, t) {
    const at = motionPointAt(clip, t);
    if (at === null) return '';
    for (const prop of motionProps(clip)) {
        const k = keyList(clip, prop).find((x) => Math.abs(x.t - at) <= EPS);
        if (k) return k.ease || '';
    }
    return '';
}

/** Set how the move into a clip's point at `t` goes. */
export function setMotionEase(doc, id, t, ease) {
    const clip = clipById(doc, id);
    const at = clip ? motionPointAt(clip, t) : null;
    if (at === null || !MOTION_EASES.includes(ease)) return doc;
    const keys = { ...(clip.keys || {}) };
    for (const prop of motionProps(clip)) {
        keys[prop] = keyList(clip, prop).map((k) => {
            if (Math.abs(k.t - at) > EPS) return k;
            const { ease: old, ...rest } = k;
            return ease ? { ...rest, ease } : rest;
        });
        if (!keys[prop].length) delete keys[prop];
    }
    return rewrite(doc, new Map([[id, withKeys(clip, keys)]]));
}

/** Take all movement off a clip: it stands still as it looks at `time` on the project's clock. */
export function clearMotion(doc, id, time = null) {
    const clip = clipById(doc, id);
    if (!clip || !isMoving(clip)) return doc;
    const t = localTime(clip, time === null ? 0 : time - clip.timing.start);
    const keys = { ...(clip.keys || {}) };
    const still = {};
    for (const prop of motionProps(clip)) {
        if (!keyList(clip, prop).length) continue;
        still[prop] = round4(motionValue(clip, prop, t));
        delete keys[prop];
    }
    return rewrite(doc, new Map([[id, withKeys({ ...clip, ...still }, keys)]]));
}

/**
 * The title gallery. Each preset is the layers it makes, in frame fractions
 * (fx, fy, fw, fh, sizeFrac); `text` is the placeholder the window replaces
 * with its own words.
 */
export const TITLE_PRESETS = {
    title: {
        seconds: 4,
        layers: [
            { type: 'text', role: 'text', fx: 0.5, fy: 0.42, sizeFrac: 0.064, weight: 800, align: 'center', color: '#ffffff', enter: 'fade', exit: 'fade' },
        ],
    },
    lowerThird: {
        seconds: 5,
        layers: [
            { type: 'shape', role: 'band', geometry: 'rect', fx: 0.24, fy: 0.82, fw: 0.42, fh: 0.11, color: 'accent', radius: 0.08, enter: 'rise', exit: 'fade' },
            { type: 'text', role: 'text', fx: 0.06, fy: 0.79, sizeFrac: 0.03, weight: 700, align: 'left', color: '#ffffff', enter: 'rise', exit: 'fade' },
        ],
    },
    endCard: {
        seconds: 5,
        layers: [
            { type: 'shape', role: 'band', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 1, fh: 1, color: '#101218', enter: 'fade', exit: 'none' },
            { type: 'text', role: 'text', fx: 0.5, fy: 0.44, sizeFrac: 0.056, weight: 800, align: 'center', color: '#ffffff', enter: 'pop', exit: 'none' },
        ],
    },
    callout: {
        seconds: 3,
        layers: [
            { type: 'text', role: 'text', fx: 0.72, fy: 0.14, sizeFrac: 0.034, weight: 700, align: 'center', color: '#ffffff', pill: 'accent', enter: 'pop', exit: 'fade' },
        ],
    },
};

/**
 * Add a title from the gallery at `at`, on the first picture lane above the
 * clips that is free for its length (a two-layer title takes two lanes, its
 * band under its words, in one group). Answers { doc, ids }.
 */
/**
 * The kinds of drawn layer Video's library adds, besides media and titles:
 * the three it began with, and the rest of what the engine draws (owner,
 * D-2026-11-20-19): a paragraph or a list, a chart, a colour adjustment, a QR
 * code, the shade along the bottom, a progress bar, the fade at the two ends
 * and an outro card.
 */
export const GRAPHIC_TYPES = ['icon', 'shape', 'wordart', 'textblock', 'chart', 'adjust', 'qr', 'gradient', 'progress', 'edges', 'outro', 'caption', 'text', 'image'];

/** Those that lie over the whole film when added: they are about all of it. */
const WHOLE_FILM = new Set(['progress', 'edges']);

/** How long an outro card lasts when added, in seconds. */
export const OUTRO_SECONDS = 4;

/**
 * What the library's overlays put on the film, before they are placed. A QR
 * code stands at its top left corner (fx, fy), as the engine draws it. It is
 * added with nothing to scan: a made-up address would be in the film, and on
 * every phone that read it, until someone noticed. The engine draws no code
 * until one is typed, and Video's stage shows where it stands meanwhile.
 */
export const OVERLAYS = {
    qr: { type: 'qr', data: '', fx: 0.76, fy: 0.08, fw: 0.18 },
    progress: { type: 'progress' },
    gradient: { type: 'gradient' },
    edges: { type: 'edges', seconds: 0.45 },
    outro: { type: 'outro', text: '' },
};

/**
 * A layer from Video's library (GRAPHIC_TYPES), at `at` for `seconds` (a
 * still's length when not said), on the lowest track above the pictures that
 * is free for all of it. It shows at once, so what was added is seen where it
 * was put; an animation in or out is the person's to give it.
 *
 * A progress bar and the fade at the two ends are added over the whole film,
 * and an outro card after its last clip, since that is what each is for; all
 * three are clips like the others, moved and trimmed on the timeline.
 */
export function addGraphic(doc, patch, { at = 0, seconds = null } = {}) {
    const type = GRAPHIC_TYPES.includes(patch?.type) ? patch.type : null;
    if (!type) return { doc, ids: [] };
    const still = Number(doc.settings?.stillSeconds) || STILL_SECONDS;
    const filmEnd = projectDuration(doc);
    const own = type === 'caption' ? captionSeconds(patch.text, { cps: doc.options?.charsPerSecond, leadIn: patch.leadIn }) : null;
    let length = Math.max(MIN_CLIP, Number(seconds) || own || (type === 'outro' ? OUTRO_SECONDS : still));
    let start = round(Math.max(0, at));
    if (WHOLE_FILM.has(type)) {
        start = 0;
        length = Math.max(length, filmEnd);
    } else if (type === 'outro') {
        start = round(filmEnd);
    }
    const timing = { start, end: round(start + length) };
    // The lowest lane above the pictures with room for it.
    let lane = 1;
    const free = (l) => laneClips(doc, 'v', l).every((p) => p.timing.end <= timing.start + EPS || p.timing.start >= timing.end - EPS);
    while (lane < 64 && !free(lane)) lane += 1;
    const [id] = freshIds(doc, type, 1);
    const none = TEXT_ANIMATIONS.none;
    const clip = { ...patch, id, type, timing, lane, enter: none, exit: none };
    // A progress bar added over the film keeps covering it as the film grows,
    // until its length is changed by hand (S12).
    if (type === 'progress' && !seconds) clip.whole = true;
    // A paragraph with its words taken out stays a clip, to be typed in again.
    if (type === 'textblock') clip.keepEmpty = true;
    delete clip.anim;
    return { doc: { ...doc, patches: [...doc.patches, clip] }, ids: [id] };
}

/*
 * The caption layer (S17, from the social app's video tool): words the person
 * types or pastes, brought in over the pictures. The engine's caption painter
 * reads every field here; each is optional and falls back to the film's own.
 */

/** How a caption's words come into view. */
export const CAPTION_REVEALS = ['type', 'word', 'line', 'karaoke', 'all'];
/** Where its panel sits, and whether it has one. */
export const CAPTION_POSITIONS = ['bottom', 'middle'];
export const CAPTION_STYLES = ['box', 'bare'];
/** Its colours: the theme's text, accent and dim, then white and black. */
export const CAPTION_COLOURS = ['text', 'accent', 'dim', '#ffffff', '#000000'];
/** The one typing speed every caption of a film shares, in characters a second. */
export const CAPTION_SPEED = { min: 4, max: 60, standard: 16 };

/** A new caption: the person's words, typed in, at the foot of the frame on its panel. */
export const NEW_CAPTION = { type: 'caption', text: '', reveal: 'type', position: 'bottom', style: 'box', color: 'text', sizeScale: 1, leadIn: 0.3 };

/** The typing speed a film's captions share. */
export function captionSpeed(doc) {
    const v = Number(doc?.options?.charsPerSecond);
    return Number.isFinite(v) && v > 0 ? v : CAPTION_SPEED.standard;
}

/** The typing speed changed, for every caption of the film at once (clamped to CAPTION_SPEED). */
export function setCaptionSpeed(doc, cps) {
    const v = Math.max(CAPTION_SPEED.min, Math.min(CAPTION_SPEED.max, Math.round(Number(cps) || CAPTION_SPEED.standard)));
    return { ...doc, options: { ...(doc.options || {}), charsPerSecond: v } };
}

/** How long a caption needs on screen: its lead-in, its words typed at `cps`, and a second and a half to read the end. */
export function captionSeconds(text, { cps = CAPTION_SPEED.standard, leadIn = 0.3 } = {}) {
    const n = String(text || '').trim().length;
    const speed = Number(cps) > 0 ? Number(cps) : CAPTION_SPEED.standard;
    const lead = Number.isFinite(Number(leadIn)) ? Number(leadIn) : 0.3;
    return round(Math.max(2, lead + n / speed + 1.5));
}

/**
 * A caption's words as the lines it splits into: the person's own line
 * breaks first; with none, its sentences (Urdu's full stop too). Never the
 * lines as wrapped on screen, which change with the frame's shape.
 */
export function captionSegments(text) {
    const s = String(text || '').trim();
    if (!s) return [];
    let parts = s.split(/\n+/).map((x) => x.trim()).filter(Boolean);
    if (parts.length < 2) {
        parts = (s.match(/[^.!?۔…]+[.!?۔…]+["')\]]*\s*|[^.!?۔…]+$/g) || [s]).map((x) => x.trim()).filter(Boolean);
    }
    return parts;
}

/**
 * One caption split into one caption a line (S17, S10): the lines follow one
 * another inside the caption's own time, each given time by its length,
 * each on a track of its own (the first on the caption's), with the
 * caption's look and a short lead-in of its own. They are one group, so
 * they move together. A caption of one line is left as it is. Answers
 * { doc, ids }.
 */
export function splitCaption(doc, id) {
    const clip = clipById(doc, id);
    if (!clip || clip.type !== 'caption' || !clip.timing) return { doc, ids: [] };
    const segs = captionSegments(clip.text);
    if (segs.length < 2) return { doc, ids: [] };
    const lead = 0.2;
    const cps = captionSpeed(doc);
    const { start, end } = clip.timing;
    const need = segs.map((s) => lead + s.length / cps + 0.5);
    const k = (end - start) / need.reduce((a, b) => a + b, 0);
    const rest = { ...doc, patches: doc.patches.filter((p) => p.id !== id) };
    const group = nextGroup(rest);
    const taken = new Set([id]);
    const ids = [];
    const made = [];
    let at = start;
    let lane = laneOf(clip);
    const look = {};
    for (const key of ['reveal', 'position', 'style', 'color', 'sizeScale']) if (clip[key] !== undefined) look[key] = clip[key];
    segs.forEach((text, i) => {
        const timing = { start: round(at), end: round(i === segs.length - 1 ? end : at + need[i] * k) };
        at += need[i] * k;
        const busy = (l) => [...rest.patches, ...made].some((p) => p.timing && sectionOf(p) === 'v' && laneOf(p) === l && p.timing.start < timing.end - EPS && p.timing.end > timing.start + EPS);
        if (i > 0) lane += 1;
        while (lane < 64 && busy(lane)) lane += 1;
        const [nid] = i === 0 ? [id] : freshIds(rest, 'caption', 1, taken);
        taken.add(nid);
        ids.push(nid);
        made.push({ ...look, id: nid, type: 'caption', text, leadIn: lead, timing, lane, group, enter: clip.enter, exit: clip.exit });
    });
    return { doc: { ...rest, patches: [...rest.patches, ...made] }, ids };
}

/**
 * The plain quick adds of the social app's video tool that need no record
 * (owner, D-2026-11-24-08): words on a pill, a QR code, the logo, a footer
 * line. A function, as the window gives the words in the reader's language;
 * the logo is the picture at `logo`. Each is a layer like any other.
 */
export function quickAdd(kind, { text = '', logo = null } = {}) {
    if (kind === 'pill') return { type: 'text', name: 'pill', text, pill: 'accent', color: '#141118', weight: 800, sizeFrac: 0.045, fx: 0.14, fy: 0.06, align: 'center' };
    if (kind === 'qr') return { ...OVERLAYS.qr };
    if (kind === 'logo') return logo ? { type: 'image', name: 'logo', url: logo, fx: 0.84, fy: 0.05, fw: 0.12, shadow: true } : null;
    if (kind === 'footer') return { type: 'text', name: 'footer', text, color: 'dim', weight: 600, sizeFrac: 0.02, fx: 0.5, fy: 0.97, align: 'center', baseline: 'bottom' };
    return null;
}

export function addTitle(doc, preset, { at = 0, text = '' } = {}) {
    const def = TITLE_PRESETS[preset] || TITLE_PRESETS.title;
    const seconds = def.seconds;
    const start = round(Math.max(0, at));
    const timing = { start, end: round(start + seconds) };
    // The lowest lane above lane 0 with room for the whole title, its layers stacked.
    const need = def.layers.length;
    let lane = 1;
    const free = (l) => laneClips(doc, 'v', l).every((p) => p.timing.end <= timing.start + EPS || p.timing.start >= timing.end - EPS);
    while (lane < 64) {
        let ok = true;
        for (let k = 0; k < need; k++) if (!free(lane + k)) ok = false;
        if (ok) break;
        lane += 1;
    }
    const group = need > 1 ? nextGroup(doc) : undefined;
    const taken = new Set();
    const ids = [];
    const patches = [...doc.patches];
    def.layers.forEach((l, k) => {
        const type = l.type;
        const [id] = freshIds(doc, type, 1, taken);
        taken.add(id);
        ids.push(id);
        const base = { id, type, name: preset, timing, lane: lane + k, enter: TEXT_ANIMATIONS[l.enter] || TEXT_ANIMATIONS.none, exit: TEXT_ANIMATIONS[l.exit] || TEXT_ANIMATIONS.none, preset, ...(group ? { group } : {}) };
        if (type === 'text') {
            // A new title's words shrink to fit the frame rather than run off it (S6).
            patches.push({ ...base, text: String(text || ''), fx: l.fx, fy: l.fy, sizeFrac: l.sizeFrac, weight: l.weight, align: l.align, color: l.color, fit: true, ...(l.pill ? { pill: l.pill } : {}) });
        } else {
            patches.push({ ...base, geometry: l.geometry, fx: l.fx, fy: l.fy, fw: l.fw, fh: l.fh, color: l.color, ...(l.radius ? { radius: l.radius } : {}) });
        }
    });
    return { doc: { ...doc, patches }, ids };
}

// ── In and Out (V10) ─────────────────────────────────────────────────────────

/** Mark the In or the Out point; the other end keeps its place (or the project's edge). */
export function setRange(doc, { in: a = null, out: b = null } = {}) {
    const end = projectDuration(doc);
    const was = doc.range || { in: 0, out: end };
    let from = a ?? was.in;
    let to = b ?? was.out;
    if (to <= from + EPS) {
        if (a !== null) to = Math.max(end, from + MIN_CLIP);
        else from = 0;
    }
    return { ...doc, range: { in: round(Math.max(0, from)), out: round(Math.max(from + MIN_CLIP, to)) } };
}

export const clearRange = (doc) => ({ ...doc, range: null });

/**
 * The project cut down to the stretch from `a` to `b`, starting at zero:
 * what Export renders when only the In to Out range is wanted. Clips that
 * cross an edge are split there, so they play exactly what they played.
 */
export function cutToRange(doc, a, b) {
    let out = splitAll(doc, b).doc;
    out = splitAll(out, a).doc;
    const keep = (p) => !p.timing || (p.timing.start < b - EPS && p.timing.end > a + EPS);
    out = { ...out, patches: out.patches.filter(keep), range: null };
    out = shift(out, clipsOf(out).map((p) => p.id), -a);
    out.markers = (out.markers || []).filter((m) => m.t >= a && m.t <= b).map((m) => ({ ...m, t: round(m.t - a) }));
    return out;
}

// ── what the renderer gets ───────────────────────────────────────────────────

/**
 * Give clips that overlap on one source their own decoder: the two halves of
 * a split clip crossfading need two elements, or both draw the same frame.
 * Answers url -> [key per clip id]; the first use of a source keeps its url.
 */
export function videoKeys(doc) {
    const keys = new Map();
    const byUrl = new Map();
    for (const p of clipsOf(doc)) {
        if (p.type !== 'video' || !p.url) continue;
        if (!byUrl.has(p.url)) byUrl.set(p.url, []);
        byUrl.get(p.url).push(p);
    }
    for (const [url, list] of byUrl) {
        list.sort((a, b) => a.timing.start - b.timing.start);
        const slotsEnd = [];
        for (const p of list) {
            let slot = slotsEnd.findIndex((end) => end <= p.timing.start + EPS);
            if (slot < 0) { slot = slotsEnd.length; slotsEnd.push(0); }
            slotsEnd[slot] = p.timing.end;
            keys.set(p.id, slot === 0 ? url : `${url}#${slot}`);
        }
    }
    return keys;
}

/**
 * The patches for `buildPlan`: pictures lane by lane from the bottom up (so a
 * lane's order is its paint order), within a lane by start, so an incoming
 * clip paints over the one it fades from; hidden picture lanes and muted
 * sound lanes switched off, and a sound lane's volume multiplied in. When a
 * lane of a section is soloed, only the soloed lanes of that section play.
 */
export function planPatches(doc) {
    const lanes = doc.lanes || { v: {}, a: {} };
    const soloed = (section) => Object.values(lanes[section] || {}).some((st) => st?.solo);
    const offBySolo = { v: soloed('v'), a: soloed('a') };
    const keys = videoKeys(doc);
    const filmEnd = projectDuration(doc);
    const pictures = [];
    const sounds = [];
    const rest = [];
    for (const p of doc.patches || []) {
        if (!p) continue;
        if (!p.timing) { rest.push(p); continue; }
        if (sectionOf(p) === 'a') {
            const state = lanes.a?.[laneOf(p)] || {};
            const vol = (p.volume ?? 1) * (state.volume ?? 1);
            sounds.push({ ...p, volume: Math.max(0, Math.min(1, vol)), visible: p.visible !== false && !state.muted && !p.muted && (!offBySolo.a || Boolean(state.solo)) });
        } else {
            const state = lanes.v?.[laneOf(p)] || {};
            const out = { ...p, visible: p.visible !== false && !state.hidden && (!offBySolo.v || Boolean(state.solo)) && !(p.type === 'text' && !String(p.text || '').trim()) };
            if (p.type === 'video' && keys.has(p.id)) out.videoKey = keys.get(p.id);
            if (p.whole && filmEnd > 0) out.timing = { start: 0, end: filmEnd };
            if (p.type === 'progress' && p.segments) out.cuts = progressCuts(doc, p.segments, out.timing);
            delete out.z;
            pictures.push(out);
        }
    }
    pictures.sort((a, b) => laneOf(a) - laneOf(b) || a.timing.start - b.timing.start);
    return [...rest, ...pictures, ...sounds];
}

/**
 * Where a progress bar in pieces breaks: at each marker, or where each clip
 * of the main track starts, inside the bar's own time (S1).
 */
function progressCuts(doc, segments, timing) {
    const at = segments === 'markers' ? (doc.markers || []).map((m) => m.t)
        : segments === 'clips' ? laneClips(doc, 'v', 0).map((p) => p.timing.start) : [];
    const inside = at.filter((t) => t > timing.start + 0.05 && t < timing.end - 0.05);
    return [...new Set(inside.map(round))].sort((a, b) => a - b);
}

/** The urls the window must load, by what loads them: pictures and videos (by key) and sounds. */
export function projectMedia(doc) {
    const keys = videoKeys(doc);
    const videos = new Map();
    const images = new Set();
    const sounds = new Set();
    for (const p of clipsOf(doc)) {
        // An end card's own logo (S3).
        if (p.type === 'outro' && p.logo) images.add(p.logo);
        if (!p.url) continue;
        if (p.type === 'video') videos.set(keys.get(p.id) || p.url, p.url);
        else if (p.type === 'image') images.add(p.url);
        else if (p.type === 'sound') sounds.add(p.url);
    }
    return { videos: [...videos].map(([key, url]) => ({ key, url })), images: [...images], sounds: [...sounds] };
}
