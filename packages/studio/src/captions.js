/**
 * Captions (docs/STUDIO.md V23 and X8): SubRip (.srt) and WebVTT (.vtt) read
 * and written, and a caption track on Video's timeline.
 *
 * A caption on the timeline is a `text` layer marked `caption`, one per cue,
 * on a picture lane of its own above everything else, so each cue is a clip:
 * moved, trimmed, split and its words edited like any title, and drawn by the
 * engine's text painter (a cue's two lines as two lines, on the scrim's pill,
 * right to left when its words are Urdu or Arabic). Burnt in is what an
 * export of the project is; kept as a track is what `captionsOf` hands back
 * to be written as a file beside it.
 *
 * Pure: text in, cues out, documents in, documents out.
 */

import { clipsOf, laneOf, sectionOf } from './project.js';
import { nextId } from './doc.js';

const round = (t) => Math.round(t * 1000) / 1000;

// ── times ───────────────────────────────────────────────────────────────────

/**
 * A cue time as seconds: `01:02:03,456` (SubRip), `01:02:03.456` or
 * `02:03.456` (WebVTT, the hours optional); a comma or a full stop before
 * the thousandths either way, as files in the wild mix them. NaN when it is
 * not a time.
 */
export function parseTime(text) {
    const m = /^\s*(?:(\d+):)?(\d{1,2}):(\d{1,2})(?:[.,](\d{1,3}))?\s*$/.exec(String(text || ''));
    if (!m) return NaN;
    const [, h, mm, ss, ms] = m;
    return (Number(h || 0) * 3600) + Number(mm) * 60 + Number(ss) + (ms ? Number(ms.padEnd(3, '0')) / 1000 : 0);
}

/** Seconds as a cue time: `00:01:02,500` for SubRip, `00:01:02.500` for WebVTT. */
export function formatTime(seconds, sep = ',') {
    const total = Math.max(0, Math.round((Number(seconds) || 0) * 1000));
    const ms = total % 1000;
    const s = Math.floor(total / 1000) % 60;
    const m = Math.floor(total / 60000) % 60;
    const h = Math.floor(total / 3600000);
    const two = (n) => String(n).padStart(2, '0');
    return `${two(h)}:${two(m)}:${two(s)}${sep}${String(ms).padStart(3, '0')}`;
}

// ── reading ─────────────────────────────────────────────────────────────────

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0', lrm: '\u200e', rlm: '\u200f' };

/** A cue's words without the markup a file may carry: <b>, <i>, <v Name>, <c.class>, inline times, {\an8}. */
function cleanWords(text) {
    return String(text)
        .replace(/<\d{1,2}:\d{2}(?::\d{2})?[.,]\d{1,3}>/g, '')
        .replace(/<\/?[a-z][^>]*>/gi, '')
        .replace(/\{\\[^}]*\}/g, '')
        .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (all, name) => {
            if (name[0] === '#') {
                const code = name[1] === 'x' || name[1] === 'X' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
                return Number.isFinite(code) ? String.fromCodePoint(code) : all;
            }
            return ENTITIES[name.toLowerCase()] ?? all;
        })
        .split('\n')
        .map((line) => line.replace(/[ \t]+$/g, ''))
        .join('\n')
        .trim();
}

const TIMING = /^\s*((?:\d+:)?\d{1,2}:\d{1,2}(?:[.,]\d{1,3})?)\s*-->\s*((?:\d+:)?\d{1,2}:\d{1,2}(?:[.,]\d{1,3})?)/;

/** The blocks of a caption file: runs of lines between blank lines, the newlines made plain. */
function blocksOf(text) {
    return String(text || '')
        .replace(/^\ufeff/, '')
        .replace(/\r\n?/g, '\n')
        .split(/\n[ \t]*\n+/)
        .map((b) => b.replace(/^\n+/, ''))
        .filter((b) => b.trim());
}

/** Cues sorted and cleaned: a cue with no words, or that ends before it starts, is dropped. */
function tidy(cues) {
    return cues
        .filter((c) => c.text && Number.isFinite(c.start) && Number.isFinite(c.end) && c.end > c.start)
        .map((c) => ({ start: round(c.start), end: round(c.end), text: c.text }))
        .sort((a, b) => a.start - b.start || a.end - b.end);
}

/** A SubRip file's cues: [{ start, end, text }], seconds, a cue's lines joined by a newline. */
export function parseSrt(text) {
    const cues = [];
    for (const block of blocksOf(text)) {
        const lines = block.split('\n');
        // The number line is optional in practice; the timing line is what a cue is.
        const at = lines.findIndex((l) => TIMING.test(l));
        if (at < 0 || at > 1) continue;
        const m = TIMING.exec(lines[at]);
        cues.push({ start: parseTime(m[1]), end: parseTime(m[2]), text: cleanWords(lines.slice(at + 1).join('\n')) });
    }
    return tidy(cues);
}

/** A WebVTT file's cues, its header, notes, styles and regions left out, and a cue's settings after its times. */
export function parseVtt(text) {
    const cues = [];
    const blocks = blocksOf(text);
    for (const [i, block] of blocks.entries()) {
        if (i === 0 && /^WEBVTT/.test(block)) {
            // The header block may run straight into the first cue without a blank line.
            const lines = block.split('\n');
            const at = lines.findIndex((l) => TIMING.test(l));
            if (at < 0) continue;
            const m = TIMING.exec(lines[at]);
            cues.push({ start: parseTime(m[1]), end: parseTime(m[2]), text: cleanWords(lines.slice(at + 1).join('\n')) });
            continue;
        }
        if (/^(NOTE|STYLE|REGION)(\s|$)/.test(block)) continue;
        const lines = block.split('\n');
        const at = lines.findIndex((l) => TIMING.test(l));
        if (at < 0 || at > 1) continue;
        const m = TIMING.exec(lines[at]);
        cues.push({ start: parseTime(m[1]), end: parseTime(m[2]), text: cleanWords(lines.slice(at + 1).join('\n')) });
    }
    return tidy(cues);
}

/**
 * A caption file of either kind, told by its header (WEBVTT) and then by its
 * name. Answers { format: 'srt' | 'vtt', cues }.
 */
export function parseCaptions(text, name = '') {
    const body = String(text || '').replace(/^\ufeff/, '');
    const vtt = /^WEBVTT/.test(body) || /\.vtt$/i.test(String(name));
    return { format: vtt ? 'vtt' : 'srt', cues: vtt ? parseVtt(body) : parseSrt(body) };
}

// ── writing ─────────────────────────────────────────────────────────────────

/** Cues as a SubRip file, numbered from one, Windows line ends as most players expect. */
export function writeSrt(cues) {
    return tidy(cues || []).map((c, i) => `${i + 1}\r\n${formatTime(c.start, ',')} --> ${formatTime(c.end, ',')}\r\n${c.text.replace(/\n/g, '\r\n')}\r\n`).join('\r\n');
}

/** Cues as a WebVTT file, with its header; the three characters WebVTT reserves escaped. */
export function writeVtt(cues) {
    const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const body = tidy(cues || []).map((c) => `${formatTime(c.start, '.')} --> ${formatTime(c.end, '.')}\n${esc(c.text)}\n`).join('\n');
    return `WEBVTT\n\n${body}`;
}

// ── on the timeline ─────────────────────────────────────────────────────────

/** Whether a layer is a caption. */
export const isCaption = (p) => Boolean(p && p.caption && p.type === 'text');

/**
 * Where captions sit in a frame of each shape: low and centred, higher in a
 * tall frame, where a phone's own buttons cover the foot of the picture.
 */
export const CAPTION_PLACE = {
    landscape: { fy: 0.9, sizeFrac: 0.03 },
    square: { fy: 0.88, sizeFrac: 0.04 },
    vertical: { fy: 0.76, sizeFrac: 0.05 },
};

/** A caption's look, as a text layer's fields, for a frame of `aspect`. */
export function captionLook(aspect = 'landscape') {
    const at = CAPTION_PLACE[aspect] || CAPTION_PLACE.landscape;
    return { fx: 0.5, fy: at.fy, sizeFrac: at.sizeFrac, weight: 600, align: 'center', baseline: 'bottom', color: '#ffffff', bg: true };
}

/** The lane captions go on: the one they are on, or the first above every other picture. */
export function captionLane(doc) {
    const caps = clipsOf(doc).filter(isCaption);
    if (caps.length) return Math.min(...caps.map(laneOf));
    let top = -1;
    for (const p of clipsOf(doc)) if (sectionOf(p) === 'v') top = Math.max(top, laneOf(p));
    return top + 1;
}

/**
 * Put cues on the timeline as captions, `at` seconds on (0 by default), on
 * the caption lane; a cue that overlaps the one before it goes on the lane
 * above, so none covers another. With `replace`, the captions already there
 * go first. Answers { doc, ids }.
 */
export function placeCaptions(doc, cues, { at = 0, replace = true, aspect = null } = {}) {
    let base = replace ? { ...doc, patches: doc.patches.filter((p) => !isCaption(p)) } : doc;
    const lane0 = captionLane(base);
    const look = captionLook(aspect || doc.options?.aspect || 'landscape');
    const ends = [];
    const taken = { patches: [...base.patches] };
    const ids = [];
    const made = [];
    for (const cue of tidy(cues || [])) {
        const start = round(cue.start + at);
        const end = round(cue.end + at);
        let k = ends.findIndex((e) => e <= start + 0.004);
        if (k < 0) { k = ends.length; ends.push(0); }
        ends[k] = end;
        const id = nextId(taken, 'text');
        taken.patches.push({ id });
        ids.push(id);
        made.push({ id, type: 'text', caption: true, name: 'caption', text: cue.text, timing: { start, end }, lane: lane0 + k, enter: { kind: 'none', seconds: 0 }, exit: { kind: 'none', seconds: 0 }, ...look });
    }
    base = { ...base, patches: [...base.patches, ...made] };
    return { doc: base, ids };
}

/** The captions on the timeline as cues, in time order: what Export Captions writes. */
export function captionsOf(doc) {
    return tidy(clipsOf(doc).filter(isCaption).filter((p) => p.visible !== false).map((p) => ({ start: p.timing.start, end: p.timing.end, text: String(p.text || '').trim() })));
}

/** Every caption laid out again for a frame of `aspect`: where it sits and how big, its words and times as they are. */
export function layoutCaptions(doc, aspect) {
    const look = captionLook(aspect);
    return { ...doc, patches: doc.patches.map((p) => (isCaption(p) ? { ...p, fy: look.fy, sizeFrac: look.sizeFrac } : p)) };
}
