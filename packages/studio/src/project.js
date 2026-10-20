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

/** How long the project runs: to the end of its last clip. */
export function projectDuration(doc) {
    let end = 0;
    for (const p of clipsOf(doc)) end = Math.max(end, p.timing.end);
    return round(end);
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

const setTiming = (p, start, end) => ({ ...p, timing: { start: round(start), end: round(end) } });

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

/** The longest a clip may run from `offset`: its source's length, when known. */
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
export function nextMarker(doc, t) {
    return (doc.markers || []).find((m) => m.t > t + EPS) || null;
}

/** The last marker before `t`, or null. */
export function previousMarker(doc, t) {
    const before = (doc.markers || []).filter((m) => m.t < t - EPS);
    return before[before.length - 1] || null;
}

// ── transitions (V15) ────────────────────────────────────────────────────────

/** The transitions, each an incoming clip's enter (and for push the outgoing one's exit). */
export const TRANSITIONS = ['fade', 'dip', 'wipe', 'slide', 'push', 'zoom'];
const ENTER_KIND = { fade: 'fade', dip: 'fade', wipe: 'wipe', slide: 'slide-left', push: 'push', zoom: 'zoom' };
const OVERLAPS = new Set(['fade', 'wipe', 'slide', 'push', 'zoom']);

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

/**
 * Put a transition between a clip and the one before it on its lane, or take
 * it off (`kind: 'none'`). A crossfade, a wipe, a slide, a push and a zoom
 * overlap the two clips by the transition's length: the incoming clip (and
 * everything after it on its lane, with their groups) moves up by that much,
 * and the sound crossfades with the picture. A dip goes to black and back
 * without an overlap. The length is held to half of each clip.
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
    const outExit = kind === 'push' ? { kind: 'push', seconds: d } : kind === 'dip' ? { kind: 'fade', seconds: ramp } : { kind: 'none', seconds: 0 };
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
    if (before) changes.set(before.id, { ...before, exit: { kind: 'none', seconds: 0 } });
    const inSound = soundOf(doc, clip);
    const outSound = soundOf(doc, before);
    if (inSound) changes.set(inSound.id, { ...inSound, enter: { kind: 'fade', seconds: 0 } });
    if (outSound) changes.set(outSound.id, { ...outSound, exit: { kind: 'fade', seconds: 0 } });
    out = rewrite(out, changes);
    if (OVERLAPS.has(kind)) {
        const moved = rippleSet(out, sectionOf(clip), laneOf(clip), clip.timing.start);
        out = shift(out, withPartners(out, moved), seconds);
    }
    return out;
}

// ── titles (V21, V22) ────────────────────────────────────────────────────────

/** Text animations, in and out: each an envelope kind. */
export const TEXT_ANIMATIONS = {
    none: { kind: 'none', seconds: 0 },
    fade: { kind: 'fade', seconds: 0.5 },
    rise: { kind: 'slide-up', seconds: 0.6, dist: 60 },
    type: { kind: 'type-on', seconds: 0 },
    pop: { kind: 'zoom', seconds: 0.45 },
};

/** Which animation an envelope is, by name. */
export function animationName(env) {
    if (!env || env.kind === 'none' || !env.kind) return 'none';
    if (env.kind === 'type-on') return 'type';
    if (env.kind === 'slide-up') return 'rise';
    if (env.kind === 'zoom') return 'pop';
    return 'fade';
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
            patches.push({ ...base, text: String(text || ''), fx: l.fx, fy: l.fy, sizeFrac: l.sizeFrac, weight: l.weight, align: l.align, color: l.color, ...(l.pill ? { pill: l.pill } : {}) });
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
            delete out.z;
            pictures.push(out);
        }
    }
    pictures.sort((a, b) => laneOf(a) - laneOf(b) || a.timing.start - b.timing.start);
    return [...rest, ...pictures, ...sounds];
}

/** The urls the window must load, by what loads them: pictures and videos (by key) and sounds. */
export function projectMedia(doc) {
    const keys = videoKeys(doc);
    const videos = new Map();
    const images = new Set();
    const sounds = new Set();
    for (const p of clipsOf(doc)) {
        if (!p.url) continue;
        if (p.type === 'video') videos.set(keys.get(p.id) || p.url, p.url);
        else if (p.type === 'image') images.add(p.url);
        else if (p.type === 'sound') sounds.add(p.url);
    }
    return { videos: [...videos].map(([key, url]) => ({ key, url })), images: [...images], sounds: [...sounds] };
}
