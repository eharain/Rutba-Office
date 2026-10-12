/**
 * Export presets by where a video is going (docs/STUDIO.md X11): the frame's
 * shape and size, the frame rate and the bitrates each destination wants,
 * and the project made ready for it.
 *
 * A project laid out for one shape exports to another without being touched:
 * every layer is in fractions of the frame, captions move to where captions
 * sit in that shape, and an introduction's scenes are laid out again for it
 * by `relayout` (intro.js hands its own in). Pure: documents in, documents out.
 */

import { layoutCaptions } from './captions.js';

/** The frame's shape by its key, as the engine's ASPECTS name them. */
const SHAPES = { landscape: { w: 16, h: 9 }, square: { w: 1, h: 1 }, vertical: { w: 9, h: 16 } };

/**
 * The destinations. `aspect` null keeps the project's own shape; `short` is
 * the frame's shorter side in pixels; `bitrate` and `audioBitrate` in bits a
 * second; `maxSeconds` where the destination takes no longer.
 */
export const EXPORT_PRESETS = [
    { id: 'project', aspect: null, short: 1080, fps: 30, bitrate: 0, audioBitrate: 128000 },
    { id: 'youtube', aspect: 'landscape', short: 1080, fps: 30, bitrate: 10_000_000, audioBitrate: 192000 },
    { id: 'social-wide', aspect: 'landscape', short: 1080, fps: 30, bitrate: 8_000_000, audioBitrate: 128000 },
    { id: 'social-square', aspect: 'square', short: 1080, fps: 30, bitrate: 6_000_000, audioBitrate: 128000 },
    { id: 'vertical', aspect: 'vertical', short: 1080, fps: 30, bitrate: 8_000_000, audioBitrate: 128000, maxSeconds: 180 },
    { id: 'email', aspect: null, short: 480, fps: 24, bitrate: 900_000, audioBitrate: 96000, maxBytes: 20 * 1024 * 1024 },
];

export const presetById = (id) => EXPORT_PRESETS.find((p) => p.id === id) || EXPORT_PRESETS[0];

const even = (n) => Math.max(2, Math.round(n / 2) * 2);

/** The frame a preset makes from a project of `aspect`: { aspect, width, height }. */
export function presetFrame(preset, projectAspect = 'landscape') {
    const p = typeof preset === 'string' ? presetById(preset) : preset || EXPORT_PRESETS[0];
    const aspect = p.aspect || (SHAPES[projectAspect] ? projectAspect : 'landscape');
    const s = SHAPES[aspect];
    const short = p.short || 1080;
    const long = (short * Math.max(s.w, s.h)) / Math.min(s.w, s.h);
    return s.w >= s.h ? { aspect, width: even(long), height: even(short) } : { aspect, width: even(short), height: even(long) };
}

/** About how big the file comes out, in bytes, for a project of `seconds`; null for the project's own quality. */
export function presetBytes(preset, seconds) {
    const p = typeof preset === 'string' ? presetById(preset) : preset;
    if (!p?.bitrate) return null;
    return Math.round(((p.bitrate + (p.audioBitrate || 0)) / 8) * Math.max(0, Number(seconds) || 0) * 1.02);
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
 * The project made ready for a preset: the frame's shape, size, rate and
 * bitrates in its options, its captions where that shape puts them, and
 * whatever `relayout(doc, aspect)` lays out again (an introduction's scenes)
 * when the shape changes. The project's own preset leaves it as it is.
 */
export function presetDoc(doc, preset, { relayout = null } = {}) {
    const p = typeof preset === 'string' ? presetById(preset) : preset || EXPORT_PRESETS[0];
    if (p.id === 'project') return doc;
    const projectAspect = doc.options?.aspect || 'landscape';
    const frame = presetFrame(p, projectAspect);
    let out = { ...doc, options: { ...doc.options, aspect: frame.aspect, frame: { width: frame.width, height: frame.height }, fps: p.fps, bitrate: p.bitrate, audioBitrate: p.audioBitrate } };
    if (frame.aspect !== projectAspect) {
        out = layoutCaptions(out, frame.aspect);
        if (typeof relayout === 'function') out = relayout(out, frame.aspect);
    }
    return out;
}
