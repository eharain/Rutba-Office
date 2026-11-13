/**
 * Spoken text layers (D-2026-11-13-31): a text layer's lines read aloud, one
 * after another, each shown while its own voice plays.
 *
 * In Video a spoken layer is, for each line, a `text` clip with the layer's
 * look and that line's words and a `sound` clip of its voice under it, the
 * two in one group, so they move, trim and split together. Every clip of the
 * layer carries `spoken: { layer, line, voice, speed }`; a title's band
 * carries `spoken: { layer, band: true }` and is stretched over all the
 * lines. Show the text and Play the voice are the clips' `visible`, which the
 * renderer, the export and Export Captions already follow.
 *
 * In the studio (Recorder and Live Broadcast) a title source carries `lines`
 * and `spoken: { show, play, voice, speed }`; the lines read while recording
 * can be handed to Video as timed text clips (`placeSpokenCues`).
 *
 * Pure: no DOM, no speech; documents in, documents out.
 */

import { clipsOf, laneClips, laneOf, EPS, MIN_CLIP } from './project.js';
import { nextId } from './doc.js';
import { captionLook } from './captions.js';

const round = (t) => Math.round(t * 1000) / 1000;

/** The quiet between two lines, and between the sentences of one line, in seconds. */
export const LINE_GAP = 0.3;
export const SENTENCE_GAP = 0.18;

/** The speeds offered: the reader's steps. */
export const SPOKEN_SPEEDS = [0.75, 1, 1.25, 1.5, 1.75, 2];

/** The most lines a layer reads, and the most characters a studio title's lines hold. */
export const MAX_LINES = 200;
export const MAX_LINES_TEXT = 4000;

// ── lines and sound ─────────────────────────────────────────────────────────

// The marks that set the direction of text and show nothing (left-to-right and
// right-to-left marks, the Arabic letter mark, embeddings and isolates):
// pasted Urdu and Arabic often carries a line of nothing else.
const DIRECTION_MARKS = new RegExp(`[${[0x200e, 0x200f, 0x061c, 0x202a, 0x202b, 0x202c, 0x202d, 0x202e, 0x2066, 0x2067, 0x2068, 0x2069].map((c) => String.fromCharCode(c)).join('')}]`, 'g');

/** A text's lines: each as typed, a line break ending it, blank lines (and lines of direction marks alone) dropped. */
export function linesOf(text) {
    return String(text ?? '')
        .split(new RegExp('\\r\\n?|\\n|\\u2028|\\u2029'))
        .map((l) => l.replace(/\s+/g, ' ').trim())
        .filter((l) => l.replace(DIRECTION_MARKS, '').trim())
        .slice(0, MAX_LINES);
}

/** A line's sentences, for a voice that speaks one sentence at a time. */
export function sentencesOfLine(line, language) {
    const s = String(line || '').trim();
    if (!s) return [];
    if (typeof Intl !== 'undefined' && Intl.Segmenter) {
        let seg;
        try { seg = new Intl.Segmenter(language || undefined, { granularity: 'sentence' }); } catch { seg = new Intl.Segmenter(undefined, { granularity: 'sentence' }); }
        return [...seg.segment(s)].map((p) => p.segment.trim()).filter(Boolean);
    }
    return (s.match(/[^.!?۔؟।。]+[.!?۔؟।。]*/g) || [s]).map((x) => x.trim()).filter(Boolean);
}

/**
 * Sentences' samples as one line: one after another with `gap` seconds of
 * quiet between. Each part is { samples, rate }; a part at another rate is
 * stretched to the first one's, so the line plays at one rate.
 */
export function joinSamples(parts, { gap = SENTENCE_GAP } = {}) {
    const list = (parts || []).filter((p) => p && p.samples && p.samples.length);
    if (!list.length) return { samples: new Float32Array(0), rate: 24000 };
    const rate = Number(list[0].rate) || 24000;
    const quiet = Math.round(gap * rate);
    const fitted = list.map((p) => resample(p.samples, Number(p.rate) || rate, rate));
    const total = fitted.reduce((n, s) => n + s.length, 0) + quiet * (fitted.length - 1);
    const out = new Float32Array(total);
    let at = 0;
    fitted.forEach((s, i) => {
        if (i) at += quiet;
        out.set(s, at);
        at += s.length;
    });
    return { samples: out, rate };
}

function resample(samples, from, to) {
    const src = samples instanceof Float32Array ? samples : Float32Array.from(samples);
    if (from === to) return src;
    const n = Math.max(1, Math.round((src.length * to) / from));
    const out = new Float32Array(n);
    for (let i = 0; i < n; i++) {
        const x = (i * from) / to;
        const k = Math.floor(x);
        const f = x - k;
        out[i] = (src[k] ?? 0) * (1 - f) + (src[k + 1] ?? src[k] ?? 0) * f;
    }
    return out;
}

/** How long samples play, in seconds. */
export const secondsOf = (samples, rate) => (samples?.length || 0) / (Number(rate) || 24000);

/** Each line's place from `at`: lines one after another with `gap` between. [{ start, end }] */
export function lineTimes(lengths, { at = 0, gap = LINE_GAP } = {}) {
    const out = [];
    let t = Math.max(0, Number(at) || 0);
    for (const len of lengths || []) {
        const l = Math.max(MIN_CLIP, Number(len) || 0);
        out.push({ start: round(t), end: round(t + l) });
        t += l + gap;
    }
    return out;
}

/**
 * A line wrapped for a band of two lines at most: as many words as `fits`
 * allows on the first, the rest on the second (which the drawing squeezes to
 * the band's width if it is still too long).
 */
export function wrapForBand(text, fits) {
    const words = String(text || '').trim().split(/\s+/).filter(Boolean);
    if (!words.length) return [];
    if (fits(words.join(' '))) return [words.join(' ')];
    let first = words[0];
    let k = 1;
    while (k < words.length && fits(`${first} ${words[k]}`)) { first = `${first} ${words[k]}`; k += 1; }
    const rest = words.slice(k).join(' ');
    return rest ? [first, rest] : [first];
}

// ── Video: spoken layers on the timeline ────────────────────────────────────

/** Whether a clip belongs to a spoken layer. */
export const isSpoken = (p) => Boolean(p?.spoken?.layer);

/** The spoken layer a clip is in, or null. */
export const spokenLayerOf = (p) => (isSpoken(p) ? String(p.spoken.layer) : null);

/** A spoken layer's clips: its lines' text and sound in line order, and its band. */
export function spokenLayer(doc, layer) {
    const mine = clipsOf(doc).filter((p) => spokenLayerOf(p) === layer);
    const lines = [];
    for (const p of mine) {
        if (p.spoken.band) continue;
        const n = Number(p.spoken.line) || 0;
        lines[n] ||= { line: n, text: null, sound: null };
        if (p.type === 'sound') lines[n].sound = p; else if (p.type === 'text') lines[n].text = p;
    }
    return {
        layer,
        clips: mine,
        lines: lines.filter(Boolean),
        bands: mine.filter((p) => p.spoken.band),
        voice: mine.find((p) => p.spoken.voice)?.spoken.voice || null,
        speed: Number(mine.find((p) => p.spoken.speed)?.spoken.speed) || 1,
    };
}

/** The spoken layers of a project, by layer id. */
export function spokenLayers(doc) {
    const ids = [...new Set(clipsOf(doc).map(spokenLayerOf).filter(Boolean))];
    return ids.map((id) => spokenLayer(doc, id));
}

/** A fresh layer id: sp1, sp2, … */
export function newSpokenLayer(doc) {
    const used = new Set(clipsOf(doc).map(spokenLayerOf).filter(Boolean));
    let n = 1;
    while (used.has(`sp${n}`)) n += 1;
    return `sp${n}`;
}

/** Whether a layer's words are shown and its voice played: { show, play }. */
export function spokenChoice(doc, layer) {
    const { lines } = spokenLayer(doc, layer);
    const texts = lines.map((l) => l.text).filter(Boolean);
    const sounds = lines.map((l) => l.sound).filter(Boolean);
    return {
        show: texts.length ? texts.some((p) => p.visible !== false) : true,
        play: sounds.length ? sounds.some((p) => p.visible !== false) : true,
    };
}

/** Show the text, play the voice, either or both: the layer's clips shown or not. */
export function setSpokenChoice(doc, layer, { show, play } = {}) {
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (spokenLayerOf(p) !== layer) return p;
            if (p.type === 'sound') return play === undefined ? p : { ...p, visible: Boolean(play) };
            return show === undefined ? p : { ...p, visible: Boolean(show) };
        }),
    };
}

/** The words of the lines of the layer a clip is in, or of the text clip alone when it is in none. */
export function wordsToRead(doc, clipId) {
    const clip = (doc.patches || []).find((p) => p.id === clipId);
    if (!clip) return [];
    const layer = spokenLayerOf(clip);
    if (!layer) return clip.type === 'text' ? linesOf(clip.text) : [];
    return spokenLayer(doc, layer).lines.map((l) => String(l.text?.text || '').trim()).filter(Boolean);
}

// What a line does not take from the words it was read from: their place on the clock and the
// timeline, their being a caption (a line runs as long as its voice, not its cue), a transition
// from the clip before, keyframes, a lock or a take's mark.
const LOOK_LEFT_OUT = new Set(['id', 'timing', 'text', 'group', 'lane', 'spoken', 'visible', 'name', 'enter', 'exit', 'offset', 'speed', 'caption', 'transition', 'keys', 'locked', 'autoZoom', 'take']);

function lookOf(p) {
    const look = {};
    for (const [k, v] of Object.entries(p || {})) if (!LOOK_LEFT_OUT.has(k)) look[k] = v;
    return look;
}

function freeFor(clips, start, end) {
    return clips.every((p) => p.timing.end <= start + EPS || p.timing.start >= end - EPS);
}

/** The lowest lane of `section` from `from` up with nothing in [start, end] but what `skip` holds. */
function laneFor(doc, section, from, start, end, skip) {
    let lane = from;
    while (lane < 64 && !freeFor(laneClips(doc, section, lane).filter((p) => !skip.has(p.id)), start, end)) lane += 1;
    return lane;
}

function groupIds(doc, count) {
    const used = new Set((doc.patches || []).map((p) => p.group).filter(Boolean));
    const out = [];
    let n = 1;
    while (out.length < count) {
        if (!used.has(`g${n}`)) out.push(`g${n}`);
        n += 1;
    }
    return out;
}

/**
 * A text clip read onto the timeline, or a spoken layer read again.
 *
 * `from`: the id of a text clip (a title's words or a caption) or of any clip
 * of a spoken layer. `lines`: [{ text, url, name, seconds, size }], each line
 * made already, its voice a file. The layer's own text and sound clips go and
 * the lines come in their place from where it started: each line's words in
 * the layer's look on one lane, its voice under it on a sound lane free for
 * all of them, each pair one group. A title's band (the other clips of the
 * text's group) is stretched over the lines. What the person chose to show
 * and play is kept. Answers { doc, layer, ids, start, end }.
 */
export function placeSpokenLines(doc, { from, lines, voice = null, speed = 1, show, play } = {}) {
    const src = (doc.patches || []).find((p) => p.id === from);
    if (!src || !src.timing) return { doc, layer: null, ids: [], start: 0, end: 0 };
    const was = spokenLayerOf(src);
    const layer = was || newSpokenLayer(doc);
    const old = was ? spokenLayer(doc, was) : null;
    const firstText = old ? (old.lines.find((l) => l.text)?.text || src) : src;
    const choice = was ? spokenChoice(doc, was) : { show: true, play: true };
    const keepShow = show ?? choice.show;
    const keepPlay = play ?? choice.play;
    // The band: a title's other clips, in its group before it was read, or marked as the layer's since.
    const bands = old ? old.bands : (src.group ? clipsOf(doc).filter((p) => p.group === src.group && p.id !== src.id && p.type !== 'text' && p.type !== 'sound') : []);
    const gone = new Set(old ? old.clips.filter((p) => !p.spoken.band).map((p) => p.id) : [src.id]);
    // From where the first line is now (it may have been moved), else where the text stood.
    const first = old?.lines[0];
    const start = round(first ? (first.text || first.sound).timing.start : src.timing.start);
    const list = (lines || []).filter((l) => l && String(l.text || '').trim() && l.url).slice(0, MAX_LINES);
    const times = lineTimes(list.map((l) => l.seconds), { at: start });
    const end = times.length ? times[times.length - 1].end : start;
    let base = { ...doc, patches: doc.patches.filter((p) => !gone.has(p.id)) };
    const skip = new Set([...gone, ...bands.map((p) => p.id)]);
    // The words on a lane free for all the lines. A title's band stays under its words: the
    // two move up together, as far as needed for both to be free, so neither covers another clip.
    const wordsFrom = Math.max(1, laneOf(firstText));
    const freeAt = (lane) => freeFor(laneClips(base, 'v', lane).filter((p) => !skip.has(p.id)), start, end);
    let shiftBy = 0;
    if (bands.length) {
        const fits = (k) => freeAt(wordsFrom + k) && bands.every((b) => freeAt(laneOf(b) + k));
        while (shiftBy < 64 && !fits(shiftBy)) shiftBy += 1;
    }
    const textLane = bands.length ? wordsFrom + shiftBy : laneFor(base, 'v', wordsFrom, start, end, skip);
    const soundLane = laneFor(base, 'a', 0, start, end, skip);
    const look = lookOf(firstText);
    const sources = { ...(base.sources || {}) };
    const taken = { patches: [...base.patches] };
    const groups = groupIds(base, list.length);
    const made = [];
    const ids = [];
    const still = { kind: 'none', seconds: 0 };
    list.forEach((l, i) => {
        const timing = times[i];
        const tid = nextId(taken, 'text');
        taken.patches.push({ id: tid });
        const sid = nextId(taken, 'sound');
        taken.patches.push({ id: sid });
        const spoken = { layer, line: i, voice, speed };
        made.push({
            ...look, id: tid, type: 'text', name: firstText.name || 'spoken', text: String(l.text).trim(), timing, lane: textLane, group: groups[i],
            enter: i === 0 ? (firstText.enter || still) : still, exit: i === list.length - 1 ? (firstText.exit || still) : still,
            spoken, ...(keepShow ? {} : { visible: false }),
        });
        made.push({
            id: sid, type: 'sound', name: l.name || '', url: l.url, offset: 0, volume: 1, timing, lane: soundLane, group: groups[i],
            enter: { kind: 'fade', seconds: 0 }, exit: { kind: 'fade', seconds: 0 }, spoken, ...(keepPlay ? {} : { visible: false }),
        });
        sources[l.url] = { name: l.name || '', kind: 'audio', duration: round(Number(l.seconds) || 0) || null, hasAudio: true, width: null, height: null, size: l.size ?? null };
        ids.push(tid, sid);
    });
    const bandIds = new Set(bands.map((p) => p.id));
    // A band keeps its own showing or not (one hidden before stays hidden), unless the words are hidden.
    const asBand = (p) => {
        const { group, ...rest } = p;
        return { ...rest, lane: laneOf(p) + shiftBy, timing: { start, end: round(Math.max(end, start + MIN_CLIP)) }, spoken: { layer, band: true }, ...(keepShow ? {} : { visible: false }) };
    };
    base = { ...base, sources, patches: [...base.patches.map((p) => (bandIds.has(p.id) ? asBand(p) : p)), ...made] };
    return { doc: base, layer, ids, start, end };
}

/**
 * A WAV's file name for a layer's line: "<layer>-<n>.wav", n from one, with
 * the reading's own mark when given, so a line read again is a new file and
 * never one a window still holds the sound of.
 */
export const spokenFileName = (layer, line, mark = '') => `${layer}-${line + 1}${mark ? `-${mark}` : ''}.wav`;

/**
 * The lines a Recorder read while recording, handed with its takes, on
 * Video's timeline from `at`: each a text clip in the caption's look, timed
 * as it was read, marked as one spoken layer. Their voice is already in the
 * take's sound, so they have no sound of their own. `cues`: [{ start, end,
 * text }] from the recording's start. Answers { doc, layer, ids }.
 */
export function placeSpokenCues(doc, cues, { at = 0, show = true } = {}) {
    const list = (cues || [])
        .filter((c) => c && String(c.text || '').trim() && Number.isFinite(c.start) && Number.isFinite(c.end) && c.end > c.start)
        .map((c) => ({ start: round(Math.max(0, c.start) + at), end: round(c.end + at), text: String(c.text).trim() }))
        .sort((a, b) => a.start - b.start);
    if (!list.length) return { doc, layer: null, ids: [] };
    // One after another they never overlap, as they were read; an end past the next start is cut there.
    for (let i = 0; i + 1 < list.length; i++) if (list[i].end > list[i + 1].start) list[i].end = Math.max(list[i].start + MIN_CLIP, list[i + 1].start);
    const layer = newSpokenLayer(doc);
    const start = list[0].start;
    const end = list[list.length - 1].end;
    // Over every picture there is while they run, on a lane of their own: no caption or clip shares it.
    let top = 0;
    for (const p of clipsOf(doc)) if (p.type !== 'sound' && p.timing.end > start + EPS && p.timing.start < end - EPS) top = Math.max(top, laneOf(p) + 1);
    const lane = laneFor(doc, 'v', Math.max(1, top), start, end, new Set());
    const look = captionLook(doc.options?.aspect || 'landscape');
    const taken = { patches: [...doc.patches] };
    const made = list.map((c, i) => {
        const id = nextId(taken, 'text');
        taken.patches.push({ id });
        return {
            ...look, id, type: 'text', name: 'spoken', text: c.text, timing: { start: c.start, end: c.end }, lane,
            enter: { kind: 'none', seconds: 0 }, exit: { kind: 'none', seconds: 0 },
            spoken: { layer, line: i, recorded: true }, ...(show ? {} : { visible: false }),
        };
    });
    return { doc: { ...doc, patches: [...doc.patches, ...made] }, layer, ids: made.map((p) => p.id) };
}

/**
 * The takes a Recorder hands over with the lines it read: the take that
 * carries the recording's sound (`carrier`) counts as sounded, since the
 * lines' voice is in it, even when nothing else was heard at Record.
 */
export function soundedTakes(takes, cues) {
    if (!(cues || []).length) return takes || [];
    return (takes || []).map((k) => (k && k.carrier && !k.hasAudio ? { ...k, hasAudio: true } : k));
}

/**
 * Whether a recording begun now will have spoken lines in its sound: a title
 * with lines to read and Play the voice on, with a voice to read them and
 * Sound on. Counted as heard when Record is pressed, so the recording keeps
 * its sound track for them.
 */
export function speaksIntoRecording(stack, { voices = false } = {}) {
    if (!voices || !stack || stack.sound !== 'on') return false;
    return (stack.items || []).some((it) => it && it.kind === 'title' && it.spoken?.play !== false && linesOf(it.lines).length > 0);
}

// ── the studio: a title's lines ─────────────────────────────────────────────

/** A title's spoken settings as stored, or anything, as settings. */
export function normaliseSpoken(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const speed = Number(r.speed);
    return {
        show: r.show === undefined ? true : Boolean(r.show),
        play: r.play === undefined ? true : Boolean(r.play),
        voice: typeof r.voice === 'string' ? r.voice.slice(0, 120) : '',
        speed: SPOKEN_SPEEDS.includes(speed) ? speed : 1,
    };
}

/** A title's lines as stored: text, at most MAX_LINES_TEXT characters. */
export const normaliseLines = (raw) => (typeof raw === 'string' ? raw.slice(0, MAX_LINES_TEXT) : '');
