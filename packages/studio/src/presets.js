/**
 * Export presets by where a video is going (docs/STUDIO.md X11): the frame's
 * shape and size, the frame rate and the bitrates each destination wants,
 * and the project made ready for it.
 *
 * A project laid out for one shape exports to another without being touched:
 * every layer is in fractions of the frame, captions move to where captions
 * sit in that shape, and an introduction's scenes are laid out again for it
 * by `relayout` (intro.js hands its own in). Pure: documents in, documents out.
 *
 * Fit under a size (a file that must stay under, say, 50 MB): the picture's
 * bitrate worked out from the project's length and the sound's share, aimed
 * about 8% under because the encoder is single pass, the written file's size
 * read back, and encoded again lower if it is still over (`encodeToFit`).
 */

import { layoutCaptions } from './captions.js';

/** The frame's shape by its key, as the engine's ASPECTS name them. */
const SHAPES = { landscape: { w: 16, h: 9 }, square: { w: 1, h: 1 }, vertical: { w: 9, h: 16 } };

/** The bits each pixel of each frame gets at a project's Quality, as the engine's plan works its bitrate out. */
export const QUALITY_BPP = { high: 0.10, medium: 0.06, low: 0.035 };

/**
 * The destinations. `aspect` null keeps the project's own shape; `short` is
 * the frame's shorter side in pixels; `bitrate` and `audioBitrate` in bits a
 * second; `maxSeconds` where the destination takes no longer. In the order
 * the Export box shows them.
 */
export const EXPORT_PRESETS = [
    { id: 'project', aspect: null, short: 1080, fps: 30, bitrate: 0, audioBitrate: 128000 },
    { id: 'youtube', aspect: 'landscape', short: 1080, fps: 30, bitrate: 10_000_000, audioBitrate: 192000 },
    { id: 'compact', aspect: null, short: 1080, fps: 30, bitrate: 4_000_000, audioBitrate: 128000 },
    { id: 'small', aspect: null, short: 720, fps: 30, bitrate: 2_000_000, audioBitrate: 128000 },
    { id: 'social-wide', aspect: 'landscape', short: 1080, fps: 30, bitrate: 8_000_000, audioBitrate: 128000 },
    { id: 'social-square', aspect: 'square', short: 1080, fps: 30, bitrate: 6_000_000, audioBitrate: 128000 },
    { id: 'vertical', aspect: 'vertical', short: 1080, fps: 30, bitrate: 8_000_000, audioBitrate: 128000, maxSeconds: 180 },
    { id: 'email', aspect: null, short: 480, fps: 24, bitrate: 900_000, audioBitrate: 96000, maxBytes: 20 * 1024 * 1024 },
];

export const presetById = (id) => EXPORT_PRESETS.find((p) => p.id === id) || EXPORT_PRESETS[0];

const even = (n) => Math.max(2, Math.round(n / 2) * 2);

/** What the file's own boxes add to the samples: the share `presetBytes` has always allowed. */
const CONTAINER = 1.02;

/** The frame a preset makes from a project of `aspect`: { aspect, width, height }. */
export function presetFrame(preset, projectAspect = 'landscape') {
    const p = typeof preset === 'string' ? presetById(preset) : preset || EXPORT_PRESETS[0];
    const aspect = p.aspect || (SHAPES[projectAspect] ? projectAspect : 'landscape');
    const s = SHAPES[aspect];
    const short = p.short || 1080;
    const long = (short * Math.max(s.w, s.h)) / Math.min(s.w, s.h);
    return s.w >= s.h ? { aspect, width: even(long), height: even(short) } : { aspect, width: even(short), height: even(long) };
}

/** The frame a project draws at itself, its own size where it has one: { aspect, width, height }. */
export function projectFrame(doc) {
    const aspect = SHAPES[doc?.options?.aspect] ? doc.options.aspect : 'landscape';
    const own = doc?.options?.frame;
    if (own && Number(own.width) > 1 && Number(own.height) > 1) return { aspect, width: even(Number(own.width)), height: even(Number(own.height)) };
    return presetFrame('project', aspect);
}

/** The picture's bitrate at a project's Quality for a frame. */
export function qualityBitrate(frame, fps, quality = 'high') {
    return Math.round(frame.width * frame.height * fps * (QUALITY_BPP[quality] ?? QUALITY_BPP.high));
}

/**
 * About how big the file comes out, in bytes, for a project of `seconds`.
 * The project's own preset has no fixed bitrate: with its `quality`, `frame`
 * and `fps` given the size is worked out from them, without them it is null.
 * `sound: false` for a project with no sound in it.
 */
export function presetBytes(preset, seconds, { quality = null, frame = null, fps = null, sound = true } = {}) {
    const p = typeof preset === 'string' ? presetById(preset) : preset;
    if (!p) return null;
    let bitrate = p.bitrate;
    if (!bitrate) {
        if (!quality || !frame) return null;
        bitrate = qualityBitrate(frame, fps || p.fps, quality);
    }
    return Math.round(((bitrate + (sound ? p.audioBitrate || 0 : 0)) / 8) * Math.max(0, Number(seconds) || 0) * CONTAINER);
}

/**
 * What a destination would refuse or a reader would mind, as codes for the
 * window to word: 'long' (over the destination's length), 'big' (over the
 * size an email carries).
 */
export function presetWarnings(preset, seconds) {
    const p = typeof preset === 'string' ? presetById(preset) : preset;
    const out = [];
    if (p?.maxSeconds && seconds > p.maxSeconds) out.push('long');
    const bytes = presetBytes(p, seconds);
    if (p?.maxBytes && bytes && bytes > p.maxBytes) out.push('big');
    return out;
}

/**
 * A megabyte as a limit is counted: a million bytes. It is the smaller of the
 * two readings, so a file under "50 MB" here is under it either way a site
 * counts.
 */
export const MEGABYTE = 1_000_000;

/** The sizes the Export box offers, in MB: TikTok's review, Bluesky, TikTok's upload. */
export const FIT_SIZES = [50, 100, 287];

/** How far under the size the single-pass encoder is aimed. */
export const FIT_AIM = 0.92;

/** How many times a file is encoded before it is handed over as it is. */
export const FIT_TRIES = 3;

/** Below this much for the picture there is nothing to encode. */
export const FIT_FLOOR = 100_000;

/** A size in the same megabytes as the limit, for the words beside it: "4.6 MB", "46 MB", "800 KB". */
export function megabytes(bytes) {
    const n = Number(bytes) || 0;
    if (n < MEGABYTE) return `${Math.max(1, Math.round(n / 1000))} KB`;
    const m = n / MEGABYTE;
    return `${m < 10 ? Number(m.toFixed(1)) : Math.round(m)} MB`;
}

/** A bitrate in bits a second, in words: "4.4 Mb/s", "850 kb/s". */
export function bitrateWords(bits) {
    const n = Number(bits) || 0;
    return n >= 1e6 ? `${(n / 1e6).toFixed(1)} Mb/s` : `${Math.round(n / 1000)} kb/s`;
}

/** A size typed in MB as bytes, or 0 where it is not a size. */
export function parseMegabytes(text) {
    const m = Number(String(text ?? '').trim().replace(',', '.'));
    return Number.isFinite(m) && m > 0 && m < 1_000_000 ? Math.round(m * MEGABYTE) : 0;
}

/**
 * The bitrates that bring a project of `seconds` to `bytes`, aimed `aim`
 * under: { video, audio, total } in bits a second, or null without a length.
 * The sound takes its share first, at the destination's own rate; where
 * that would be more than a quarter of the whole it drops to 96 kb/s (not
 * every AAC encoder will go lower).
 * With no sound the picture has it all.
 */
export function fitBitrates({ bytes, seconds, audioBitrate = 128000, sound = true, aim = FIT_AIM }) {
    const s = Number(seconds) || 0;
    if (!(bytes > 0) || s <= 0) return null;
    const total = (bytes * aim * 8) / (s * CONTAINER);
    const audio = !sound ? 0 : audioBitrate * 4 > total ? Math.min(96000, audioBitrate) : audioBitrate;
    return { video: Math.max(0, Math.floor(total - audio)), audio, total: Math.floor(total) };
}

/** The picture's rate below which a frame looks poor: about 1 Mb/s at 1080p, in proportion to its pixels elsewhere. */
export const poorBelow = (frame) => Math.round(1_000_000 * ((frame.width * frame.height) / (1920 * 1080)));

/**
 * What fitting a preset under `bytes` comes to, for the Export box and the
 * export: { limit, video, audio, estimate, capped, poor, suggest, tooSmall,
 * short, frame, fps } or null when there is no size or no length.
 *
 * `capped` the preset's own bitrate is already under the size, so the preset
 * is what it is, the size asking nothing of it; `poor` the picture would be
 * too thin to watch at this frame, with `suggest` the shorter side to try
 * instead (720, then 480); `tooSmall` nothing is left for the picture at all.
 * `short` asks for a smaller frame than the preset's (a "Use 720 instead").
 * For the project's own preset pass its real `frame`, `fps` and `quality`.
 */
export function fitPlan(preset, seconds, { bytes, quality = 'high', frame = null, fps = null, aspect = 'landscape', sound = true, short = 0 } = {}) {
    const p = typeof preset === 'string' ? presetById(preset) : preset || EXPORT_PRESETS[0];
    const base = frame || presetFrame(p, aspect);
    const baseShort = Math.min(base.width, base.height);
    const smaller = short > 0 && short < baseShort;
    const shaped = smaller ? presetFrame({ ...p, short }, base.aspect || aspect) : base;
    const rate = fps || p.fps;
    const own = p.bitrate
        ? Math.round(p.bitrate * ((shaped.width * shaped.height) / (base.width * base.height)))
        : qualityBitrate(shaped, rate, quality);
    const sums = fitBitrates({ bytes, seconds, audioBitrate: p.audioBitrate, sound });
    if (!sums) return null;
    const capped = sums.video >= own;
    const video = capped ? own : sums.video;
    const poor = !capped && video >= FIT_FLOOR && video < poorBelow(shaped);
    const shortest = Math.min(shaped.width, shaped.height);
    return {
        limit: bytes,
        video,
        audio: sums.audio,
        estimate: Math.round(((video + sums.audio) / 8) * seconds * CONTAINER),
        capped,
        poor,
        suggest: poor ? (shortest > 720 ? 720 : shortest > 480 ? 480 : null) : null,
        tooSmall: video < FIT_FLOOR,
        short: smaller ? short : 0,
        frame: shaped,
        fps: rate,
    };
}

/**
 * Encode until the written file is under the limit. `encode(video, attempt)`
 * writes the file at that picture bitrate and answers at least { size }, the
 * size as read back from what was written. If it is over, the next go is
 * lower by the proportion it missed by, and a little more, and `onAgain({
 * attempt, size, limit, video })` says so first. After `tries` goes, at the
 * floor, or as soon as asking for less stops making the file smaller (an
 * encoder has a worst quality it will go to, and a detailed picture can reach
 * it, so that a lower bitrate changes nothing), what there is is handed over
 * and `over` says it is still too big. Answers { out, tries: [{ video, size }], over }.
 */
export async function encodeToFit({ fit, encode, onAgain = null, tries = FIT_TRIES, aim = FIT_AIM, first = null }) {
    const log = [];
    let video = first ?? fit.video;
    for (let attempt = 1; ; attempt++) {
        const out = await encode(video, attempt);
        log.push({ video, size: out.size });
        if (out.size <= fit.limit) return { out, tries: log, over: false };
        const before = log[log.length - 2];
        if (attempt >= tries || video <= FIT_FLOOR || (before && out.size >= before.size * 0.97)) return { out, tries: log, over: true };
        const scale = (fit.limit * aim) / out.size;
        const whole = (video + (fit.audio || 0)) * scale * 0.96;
        const next = Math.max(FIT_FLOOR, Math.min(Math.round(whole - (fit.audio || 0)), Math.round(video * 0.9)));
        await onAgain?.({ attempt, size: out.size, limit: fit.limit, video: next });
        video = next;
    }
}

/**
 * The project made ready for a preset: the frame's shape, size, rate and
 * bitrates in its options, its captions where that shape puts them, and
 * whatever `relayout(doc, aspect)` lays out again (an introduction's scenes)
 * when the shape changes. The project's own preset leaves it as it is.
 *
 * `fit` is what `fitPlan` answered: its bitrates in place of the preset's, on
 * the project's own preset too, and its smaller frame where it has one.
 */
export function presetDoc(doc, preset, { relayout = null, fit = null } = {}) {
    let p = typeof preset === 'string' ? presetById(preset) : preset || EXPORT_PRESETS[0];
    if (p.id === 'project') {
        if (!fit) return doc;
        const kept = { ...doc.options, bitrate: fit.video, audioBitrate: fit.audio || p.audioBitrate };
        if (fit.short && fit.frame) kept.frame = { width: fit.frame.width, height: fit.frame.height };
        return { ...doc, options: kept };
    }
    if (fit) p = { ...p, short: fit.short || p.short, bitrate: fit.video, audioBitrate: fit.audio || p.audioBitrate };
    const projectAspect = doc.options?.aspect || 'landscape';
    const frame = presetFrame(p, projectAspect);
    let out = { ...doc, options: { ...doc.options, aspect: frame.aspect, frame: { width: frame.width, height: frame.height }, fps: p.fps, bitrate: p.bitrate, audioBitrate: p.audioBitrate } };
    if (frame.aspect !== projectAspect) {
        out = layoutCaptions(out, frame.aspect);
        if (typeof relayout === 'function') out = relayout(out, frame.aspect);
    }
    return out;
}
