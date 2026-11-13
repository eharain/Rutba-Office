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
import { placeCaptions } from './captions.js';

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

/** A text's lines: each as typed, a line break ending it, blank lines dropped. */
export function linesOf(text) {
    return String(text ?? '')
        .split(new RegExp('\\r\\n?|\\n|\\u2028|\\u2029'))
        .map((l) => l.replace(/\s+/g, ' ').trim())
        .filter(Boolean)
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

const LOOK_LEFT_OUT = new Set(['id', 'timing', 'text', 'group', 'lane', 'spoken', 'visible', 'name', 'enter', 'exit', 'offset', 'speed']);

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
    const textLane = laneFor(base, 'v', Math.max(1, laneOf(firstText)), start, end, skip);
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
    const asBand = (p) => {
        const { group, visible, ...rest } = p;
        return { ...rest, timing: { start, end: round(Math.max(end, start + MIN_CLIP)) }, spoken: { layer, band: true }, ...(keepShow ? {} : { visible: false }) };
    };
    base = { ...base, sources, patches: [...base.patches.map((p) => (bandIds.has(p.id) ? asBand(p) : p)), ...made] };
    return { doc: base, layer, ids, start, end };
}

/** A WAV's file name for a layer's line: "<layer>-<n>.wav", n from one. */
export const spokenFileName = (layer, line) => `${layer}-${line + 1}.wav`;

/**
 * The lines a Recorder read while recording, handed with its takes, on
 * Video's timeline from `at`: each a text clip in the caption's look, timed
 * as it was read, marked as one spoken layer. Their voice is already in the
 * take's sound, so they have no sound of their own. `cues`: [{ start, end,
 * text }] from the recording's start. Answers { doc, layer, ids }.
 */
export function placeSpokenCues(doc, cues, { at = 0, show = true } = {}) {
    const list = (cues || []).filter((c) => c && String(c.text || '').trim() && Number.isFinite(c.start) && Number.isFinite(c.end) && c.end > c.start);
    if (!list.length) return { doc, layer: null, ids: [] };
    const layer = newSpokenLayer(doc);
    const placed = placeCaptions(doc, list, { at, replace: false });
    const mine = new Set(placed.ids);
    let n = 0;
    const order = new Map(placed.doc.patches.filter((p) => mine.has(p.id)).sort((a, b) => a.timing.start - b.timing.start).map((p) => [p.id, n++]));
    const next = {
        ...placed.doc,
        patches: placed.doc.patches.map((p) => {
            if (!mine.has(p.id)) return p;
            const { caption, ...rest } = p;
            return { ...rest, name: 'spoken', spoken: { layer, line: order.get(p.id), recorded: true }, ...(show ? {} : { visible: false }) };
        }),
    };
    return { doc: next, layer, ids: placed.ids };
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
