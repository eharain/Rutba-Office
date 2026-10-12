/**
 * The screen recorder's arithmetic (docs/STUDIO.md section 8, Phase 2).
 *
 * A recording is two files on one clock (decision D6): the screen (a whole
 * screen, one window, or a screen with a region kept by a crop) and, when it
 * is on, the camera. The microphone goes into the camera's file and the
 * computer's own sound into the screen's, so each sound lands on a lane of
 * its own; without a camera both sounds go into the screen's file, mixed.
 *
 * Everything here is pure: the settings and what they cost on disk, the
 * camera's inset in the frame, a drawn region as a crop, the clock that
 * pauses, the writer that streams a recorder's pieces to a file a few
 * megabytes at a time, and where the takes land on Video's timeline. The
 * window that records (apps/desktop/renderer/apps/video/recorder/) does the
 * media; a web host would do the same with its own transport.
 */

import { addMedia, clipsOf, sectionOf, laneOf, projectDuration, clipById } from './project.js';

// ── settings ────────────────────────────────────────────────────────────────

/** The frame rates offered. */
export const FRAME_RATES = [15, 30, 60];

/** The largest the screen is recorded at, by its shorter side; 0 keeps the source's own size. */
export const SIZE_CAPS = [720, 1080, 1440, 0];

/**
 * Quality as bits per pixel of each frame: what sets the bitrate for a given
 * size and frame rate. Standard keeps a 1080p screen at 30 frames a second
 * near 6 Mbps, under the 8 the plan asks of the defaults.
 */
export const QUALITIES = { small: 0.06, standard: 0.1, high: 0.14 };

/** The camera's inset, as a share of the frame's width. */
export const INSET_SIZES = { small: 0.18, medium: 0.25, large: 0.34 };

/** The corners the camera may sit in. */
export const CORNERS = ['tl', 'tr', 'bl', 'br'];

/** The shapes the camera may be cut to. */
export const CAMERA_SHAPES = ['circle', 'rounded', 'rect'];

/** The sound recorded with each file, in bits a second. */
export const AUDIO_BITRATE = 128000;

/** Below this many minutes of disk at the rate being recorded, the recorder warns. */
export const LOW_DISK_MINUTES = 5;

/** Below this many seconds, it stops on its own and keeps what it has. */
export const STOP_DISK_SECONDS = 20;

/** What a recorder starts with. */
export const RECORDER_DEFAULTS = Object.freeze({
    source: 'screen', // 'screen' | 'window' | 'region'
    fps: 30,
    cap: 1080,
    quality: 'standard',
    countdown: 3,
    camera: true,
    cameraId: '',
    cameraHeight: 720,
    mirror: true,
    corner: 'br',
    insetSize: 'medium',
    shape: 'circle',
    mic: true,
    micId: '',
    system: false,
    minimise: true,
    folder: '',
});

const pick = (v, list, fallback) => (list.includes(v) ? v : fallback);

/** Stored settings, or anything, as settings the recorder can use. */
export function normaliseSettings(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const d = RECORDER_DEFAULTS;
    return {
        source: pick(r.source, ['screen', 'window', 'region'], d.source),
        fps: pick(Number(r.fps), FRAME_RATES, d.fps),
        cap: pick(Number(r.cap), SIZE_CAPS, d.cap),
        quality: pick(r.quality, Object.keys(QUALITIES), d.quality),
        countdown: pick(Number(r.countdown), [0, 3, 5, 10], d.countdown),
        camera: r.camera === undefined ? d.camera : Boolean(r.camera),
        cameraId: typeof r.cameraId === 'string' ? r.cameraId : '',
        cameraHeight: pick(Number(r.cameraHeight), [480, 720, 1080], d.cameraHeight),
        mirror: r.mirror === undefined ? d.mirror : Boolean(r.mirror),
        corner: pick(r.corner, CORNERS, d.corner),
        insetSize: pick(r.insetSize, Object.keys(INSET_SIZES), d.insetSize),
        shape: pick(r.shape, CAMERA_SHAPES, d.shape),
        mic: r.mic === undefined ? d.mic : Boolean(r.mic),
        micId: typeof r.micId === 'string' ? r.micId : '',
        system: Boolean(r.system),
        minimise: r.minimise === undefined ? d.minimise : Boolean(r.minimise),
        folder: typeof r.folder === 'string' ? r.folder : '',
    };
}

const even = (n) => Math.max(2, Math.round(n / 2) * 2);

/** A source's size held to `cap` on its shorter side, keeping its shape; even numbers, as encoders want. */
export function capSize(width, height, cap) {
    const w = Math.max(2, Number(width) || 0);
    const h = Math.max(2, Number(height) || 0);
    const short = Math.min(w, h);
    const k = cap && short > cap ? cap / short : 1;
    return { width: even(w * k), height: even(h * k) };
}

/** The video bitrate for a size, a frame rate and a quality, in bits a second. */
export function videoBitrate({ width, height, fps = 30, quality = 'standard' }) {
    const bpp = QUALITIES[quality] ?? QUALITIES.standard;
    const raw = Math.max(1, width) * Math.max(1, height) * Math.max(1, fps) * bpp;
    return Math.round(Math.min(40e6, Math.max(250e3, raw)));
}

/**
 * What a recording costs the disk, in bytes a second: the screen, the camera
 * when it is on, a sound track with each file that has one, and a little for
 * the container.
 */
export function bytesPerSecond(settings, { screen = { width: 1920, height: 1080 }, camera = null } = {}) {
    const s = normaliseSettings(settings);
    const shot = capSize(screen.width, screen.height, s.cap);
    let bits = videoBitrate({ ...shot, fps: s.fps, quality: s.quality });
    const withCamera = s.camera && camera;
    if (withCamera) bits += videoBitrate({ width: camera.width, height: camera.height, fps: 30, quality: s.quality });
    const screenSound = s.system || (s.mic && !withCamera);
    const cameraSound = withCamera && s.mic;
    bits += (screenSound ? AUDIO_BITRATE : 0) + (cameraSound ? AUDIO_BITRATE : 0);
    return Math.round((bits / 8) * 1.03);
}

/** Minutes of recording the free space holds at that rate. */
export function minutesLeft(freeBytes, rate) {
    if (!(rate > 0)) return Infinity;
    return Math.max(0, Number(freeBytes) || 0) / rate / 60;
}

/** Whether there is room: 'ok', 'low' (under five minutes) or 'stop' (under twenty seconds). */
export function diskState(freeBytes, rate) {
    const minutes = minutesLeft(freeBytes, rate);
    const level = minutes * 60 < STOP_DISK_SECONDS ? 'stop' : minutes < LOW_DISK_MINUTES ? 'low' : 'ok';
    return { minutes, level };
}

// ── where things sit ────────────────────────────────────────────────────────

/** The space between the camera's inset and the frame's edge, as a share of the frame's width. */
export const INSET_MARGIN = 0.025;

const round4 = (n) => Math.round(n * 10000) / 10000;

/**
 * The camera's inset in a frame of W by H: { fx, fy, fw, fh }, the top left
 * corner and the size as shares of the frame. A circle is as tall as it is
 * wide on screen; a rectangle takes the camera's own shape (`aspect`, its
 * height over its width).
 */
export function insetRect({ corner = 'br', size = 'medium', shape = 'circle', W = 1920, H = 1080, aspect = 9 / 16 } = {}) {
    const fw = INSET_SIZES[size] ?? (Number(size) > 0 ? Math.min(0.9, Number(size)) : INSET_SIZES.medium);
    const px = fw * W;
    const ph = shape === 'circle' ? px : px * (aspect > 0 ? aspect : 9 / 16);
    const fh = Math.min(0.95, ph / H);
    const mx = INSET_MARGIN;
    const my = (INSET_MARGIN * W) / H;
    const right = corner === 'tr' || corner === 'br';
    const bottom = corner === 'bl' || corner === 'br';
    return {
        fx: round4(right ? 1 - mx - fw : mx),
        fy: round4(bottom ? 1 - my - fh : my),
        fw: round4(fw),
        fh: round4(fh),
    };
}

/** The corner nearest an inset's middle: where one dropped after a drag settles. */
export function cornerOf({ fx = 0, fy = 0, fw = 0, fh = 0 } = {}) {
    const cx = fx + fw / 2;
    const cy = fy + fh / 2;
    return `${cy < 0.5 ? 't' : 'b'}${cx < 0.5 ? 'l' : 'r'}`;
}

/**
 * A region drawn on a screen, as the crop the screen's layer keeps:
 * fractions of the screen, held inside it. `rect` is in any units with
 * `size` in the same ones. A region too small to be meant (under 2% either
 * way) is no region; the whole screen is none either.
 */
export function regionCrop(rect, size = { width: 1, height: 1 }) {
    if (!rect) return null;
    const W = Math.max(1e-9, Number(size.width) || 1);
    const H = Math.max(1e-9, Number(size.height) || 1);
    let x0 = Math.min(rect.x, rect.x + rect.w) / W;
    let y0 = Math.min(rect.y, rect.y + rect.h) / H;
    let x1 = Math.max(rect.x, rect.x + rect.w) / W;
    let y1 = Math.max(rect.y, rect.y + rect.h) / H;
    x0 = Math.max(0, Math.min(1, x0));
    y0 = Math.max(0, Math.min(1, y0));
    x1 = Math.max(0, Math.min(1, x1));
    y1 = Math.max(0, Math.min(1, y1));
    const w = x1 - x0;
    const h = y1 - y0;
    if (w < 0.02 || h < 0.02) return null;
    if (x0 < 0.001 && y0 < 0.001 && w > 0.998 && h > 0.998) return null;
    return { x: round4(x0), y: round4(y0), w: round4(w), h: round4(h) };
}

// ── the clock ───────────────────────────────────────────────────────────────

/**
 * The recording's clock: how long has been recorded, the time spent paused
 * left out. Every call takes the time it is (milliseconds), so it is the same
 * clock in a test as in a window.
 */
export function createClock() {
    let started = null;
    let heldAt = null;
    let held = 0;
    return {
        start(now) { started = now; heldAt = null; held = 0; },
        pause(now) { if (started != null && heldAt == null) heldAt = now; },
        resume(now) { if (heldAt != null) { held += now - heldAt; heldAt = null; } },
        get running() { return started != null; },
        get paused() { return heldAt != null; },
        /** Milliseconds recorded by `now`. */
        elapsed(now) {
            if (started == null) return 0;
            const end = heldAt != null ? heldAt : now;
            return Math.max(0, end - started - held);
        },
    };
}

// ── streaming to a file ─────────────────────────────────────────────────────

/** The most of a recording one message carries to its file. */
export const SLICE = 4 * 1024 * 1024;

/**
 * Writes a recorder's pieces to a file in the order they came, `slice` bytes
 * a message at most, so the window holds none of the recording and no one
 * message carries much of it. `append(bytes)` writes (answers a promise); a
 * piece is anything with `size`, `slice(a, b)` and `arrayBuffer()` (a Blob).
 * After a failure nothing more is written and `done()` says why.
 */
export function createChunkWriter({ append, slice = SLICE }) {
    let queue = Promise.resolve();
    let written = 0;
    let failed = null;
    const put = async (part) => {
        for (let at = 0; at < part.size && !failed; at += slice) {
            const bytes = new Uint8Array(await part.slice(at, Math.min(part.size, at + slice)).arrayBuffer());
            await append(bytes);
            written += bytes.length;
        }
    };
    return {
        push(part) {
            if (!part || !part.size) return queue;
            queue = queue.then(() => (failed ? null : put(part))).catch((err) => { failed = err || new Error('write failed'); });
            return queue;
        },
        get written() { return written; },
        get failed() { return failed; },
        async done() {
            await queue;
            return { written, failed };
        },
    };
}

/** A file's name for a take: its stem, the date and time it began, its extension. */
export function takeFileName(stem, date = new Date(), ext = 'webm') {
    const pad = (n) => String(n).padStart(2, '0');
    const when = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}.${pad(date.getMinutes())}.${pad(date.getSeconds())}`;
    const clean = String(stem || 'Recording').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Recording';
    return `${clean} ${when}.${String(ext).replace(/^\./, '')}`;
}

// ── onto the timeline (R14) ─────────────────────────────────────────────────

const EPS = 0.004;

/** Whether nothing on a lane overlaps start to end. */
function laneFree(doc, section, lane, start, end) {
    return !clipsOf(doc).some((p) => sectionOf(p) === section && laneOf(p) === lane && p.timing.start < end - EPS && p.timing.end > start + EPS);
}

/** The lowest lane at or above `from` with room from start to end. */
function freeLane(doc, section, from, start, end, taken = new Set()) {
    let lane = Math.max(0, from);
    while (taken.has(lane) || !laneFree(doc, section, lane, start, end)) lane += 1;
    return lane;
}

/**
 * Put a recording's takes on a project's timeline: the screen on the picture
 * lane (the first free one, the main lane in a new project), the camera on the
 * lane above it as an inset, each file's sound on a sound lane of its own,
 * all of it in one group, as it was recorded on one clock. They go at `at`,
 * or after everything already there.
 *
 * `takes`: [{ role: 'screen' | 'camera', url, name, duration, width, height,
 * hasAudio, size, crop?, inset?: { fx, fy, fw, fh }, shape?, mirror?,
 * sound?: 'voice' | 'computer' }].
 * `unfinished` marks a project written while still recording, for the
 * window that recovers it to settle the lengths (settleTakes).
 * Answers { doc, ids, group }.
 */
export function placeTakes(doc, takes, { at = null, unfinished = false } = {}) {
    const list = (takes || []).filter((tk) => tk && tk.url && (tk.role === 'screen' || tk.role === 'camera'));
    if (!list.length) return { doc, ids: [], group: null };
    const start = at ?? projectDuration(doc);
    const order = [...list].sort((a, b) => (a.role === b.role ? 0 : a.role === 'screen' ? -1 : 1));
    let next = doc;
    const ids = [];
    let group = null;
    let screenLane = -1;
    const usedSound = new Set();
    for (const take of order) {
        const length = Math.max(0.1, Number(take.duration) || 0.1);
        const end = start + length;
        const from = take.role === 'camera' ? screenLane + 1 : 0;
        const lane = freeLane(next, 'v', from, start, end);
        if (take.role === 'screen') screenLane = lane;
        // The sound's lane is found before the take is added, or the take's own sound would be in the way.
        const soundLane = take.hasAudio ? freeLane(next, 'a', 0, start, end, usedSound) : null;
        const r = addMedia(next, { url: take.url, name: take.name, kind: 'video', duration: length, hasAudio: Boolean(take.hasAudio), width: take.width, height: take.height, size: take.size ?? null }, { at: start, lane });
        next = r.doc;
        const [picId, soundId] = r.ids;
        const made = new Set(r.ids);
        const fresh = clipById(next, picId);
        group = group || fresh.group || `rec-${picId}`;
        if (soundLane != null) usedSound.add(soundLane);
        next = {
            ...next,
            patches: next.patches.map((p) => {
                if (!made.has(p.id)) return p;
                const base = { ...p, group, take: p.type === 'sound' ? (take.sound || (take.role === 'camera' ? 'voice' : 'computer')) : take.role };
                if (p.type === 'sound') return { ...base, lane: soundLane };
                if (take.role === 'screen') {
                    const out = { ...base, fit: 'contain' };
                    if (take.crop) out.crop = take.crop;
                    return out;
                }
                const inset = take.inset || insetRect({ W: 1920, H: 1080, aspect: take.width && take.height ? take.height / take.width : 9 / 16, shape: take.shape || 'circle' });
                return { ...base, fx: inset.fx, fy: inset.fy, fw: inset.fw, fh: inset.fh, insetShape: take.shape || 'circle', flipX: take.mirror !== false };
            }),
        };
        ids.push(...r.ids);
    }
    if (unfinished) next = { ...next, recording: { unfinished: true, group } };
    return { doc: next, ids, group };
}

/**
 * A project recovered from a recording that never stopped: its takes were
 * written down with the length they had at the last copy, and the files ran
 * on after it. `lengths` maps a take's url to its file's real length; each
 * clip of the take is given it, and the project is a finished one.
 */
export function settleTakes(doc, lengths = {}) {
    if (!doc?.recording?.unfinished) return doc;
    const group = doc.recording.group;
    const patches = doc.patches.map((p) => {
        if (!p.timing || !p.take || (group && p.group !== group)) return p;
        const real = Number(lengths[p.url]);
        if (!(real > 0)) return p;
        const end = Math.round((p.timing.start + Math.max(0.1, real - (p.offset || 0))) * 1000) / 1000;
        return { ...p, timing: { ...p.timing, end } };
    });
    const sources = { ...(doc.sources || {}) };
    for (const [url, real] of Object.entries(lengths)) if (sources[url] && real > 0) sources[url] = { ...sources[url], duration: real };
    const { recording, ...rest } = doc;
    return { ...rest, sources, patches };
}
