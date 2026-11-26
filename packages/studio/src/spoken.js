/**
 * Spoken text layers (D-2026-11-13-31): a text layer's lines read aloud, one
 * after another, each shown while its own voice plays.
 *
 * In Video a spoken layer is, for each line, a `text` clip with the layer's
 * look and that line's words and a `sound` clip of its voice under it, the
 * two in one group, so they move, trim and split together. Every clip of the
 * layer carries `spoken: { layer, line, voice, speed }`; a title's band
 * carries `spoken: { layer, band: true }` and is stretched over all the
 * lines. The one choice of three (Show and read aloud, Read aloud only, Show
 * only: SPEAK_CHOICES) is stored as the clips' `visible`, the text clips' for
 * showing and the sound clips' for reading, which the renderer, the export and
 * Export Captions already follow.
 *
 * In the studio (Recorder and Live Broadcast) a title source carries `lines`
 * and `spoken: { show, play, voice, speed }`; the lines read while recording
 * can be handed to Video as timed text clips (`placeSpokenCues`).
 *
 * Pure: no DOM, no speech; documents in, documents out.
 */

import { clipsOf, laneClips, laneOf, sectionOf, EPS, MIN_CLIP, NEW_CAPTION, captionSegments } from './project.js';
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
        if (p.type === 'sound') lines[n].sound = p; else if (p.type === 'text' || p.type === 'caption') lines[n].text = p;
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

/**
 * Whether a layer's words are shown and its voice played: { show, play }.
 * Words typed to be shown only, with no voice made (placeTypedLines), have
 * no voice to play until they are read again with one.
 */
export function spokenChoice(doc, layer) {
    const { lines } = spokenLayer(doc, layer);
    const texts = lines.map((l) => l.text).filter(Boolean);
    const sounds = lines.map((l) => l.sound).filter(Boolean);
    return {
        show: texts.length ? texts.some((p) => p.visible !== false) : true,
        play: sounds.length ? sounds.some((p) => p.visible !== false) : !texts.some((p) => p.spoken?.shown),
    };
}

/** How a read caption line's typing goes with its voice: with it, or the voice a moment first (S19). */
export const TYPING_TIMINGS = ['with', 'first'];
/** How far the voice is ahead of the typing with 'first', in seconds. */
export const VOICE_FIRST_SECONDS = 0.4;
/** How long a read caption line stands whole at the end of its voice, in seconds. */
export const TYPED_HOLD_SECONDS = 0.1;

/**
 * How a caption line read aloud types (S19): with its voice played, over the
 * voice's length, so the last letter lands as the voice ends, starting with
 * it ('with') or a moment after it ('first'); with no voice played, at the
 * film's own typing speed after the caption's usual lead-in. Answers the
 * caption's { leadIn, pace }; a pace of undefined is the film's own speed.
 */
export function typingWithVoice(clip, { play = true, typing = 'with' } = {}) {
    if (!play || !clip?.timing) return { leadIn: NEW_CAPTION.leadIn, pace: undefined };
    const length = Math.max(0, clip.timing.end - clip.timing.start);
    const lead = typing === 'first' ? Math.min(VOICE_FIRST_SECONDS, length / 2) : 0;
    // The last letter is down a moment before the line goes, as the voice ends.
    const hold = Math.min(TYPED_HOLD_SECONDS, length / 4);
    return { leadIn: lead, pace: Math.round(Math.max(0.1, length - lead - hold) * 1000) / 1000 };
}

/** A caption line of a spoken layer as it types: in step with its voice when `play`. */
function typedLine(p, play) {
    return p.type === 'caption' ? { ...p, ...typingWithVoice(p, { play, typing: p.spoken?.typing }) } : p;
}

/** A choice of the three (speakSwitches gives its { show, play }): the layer's text clips shown or not, its sound clips heard or not. */
export function setSpokenChoice(doc, layer, { show, play } = {}) {
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (spokenLayerOf(p) !== layer) return p;
            if (p.type === 'sound') return play === undefined ? p : { ...p, visible: Boolean(play) };
            // A caption read aloud types with its voice while the voice plays, and at its own speed without it.
            const typed = play === undefined || p.spoken?.band ? p : typedLine(p, Boolean(play));
            return show === undefined ? typed : { ...typed, visible: Boolean(show) };
        }),
    };
}

/** The layer's caption lines typed with their voices, or the voices a moment first: `typing` one of TYPING_TIMINGS. */
export function setSpokenTyping(doc, layer, typing) {
    const how = TYPING_TIMINGS.includes(typing) ? typing : 'with';
    const { play } = spokenChoice(doc, layer);
    return {
        ...doc,
        patches: doc.patches.map((p) => (spokenLayerOf(p) === layer && p.type === 'caption' && !p.spoken.band
            ? typedLine({ ...p, spoken: { ...p.spoken, typing: how } }, play)
            : p)),
    };
}

/** How a layer's caption lines type with their voices, or null when it has none. */
export function spokenTypingOf(doc, layer) {
    const line = spokenLayer(doc, layer).lines.map((l) => l.text).find((p) => p?.type === 'caption');
    return line ? (line.spoken.typing === 'first' ? 'first' : 'with') : null;
}

/** The words of the lines of the layer a clip is in, or of the text clip alone when it is in none. */
export function wordsToRead(doc, clipId) {
    const clip = (doc.patches || []).find((p) => p.id === clipId);
    if (!clip) return [];
    const layer = spokenLayerOf(clip);
    // A caption reads as it splits: at the person's line breaks, else at its sentences.
    if (!layer) return clip.type === 'text' ? linesOf(clip.text) : clip.type === 'caption' ? captionSegments(clip.text).slice(0, MAX_LINES) : [];
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

/** Whether a track is locked: nothing is put onto it or moved onto it. */
const lockedLane = (doc, section, lane) => Boolean(doc.lanes?.[section]?.[lane]?.locked);

/** Whether a lane of `section` is unlocked and holds nothing in [start, end] but what `skip` holds. */
function laneFree(doc, section, lane, start, end, skip) {
    return !lockedLane(doc, section, lane) && freeFor(laneClips(doc, section, lane).filter((p) => !skip.has(p.id)), start, end);
}

/** The first lane of `section` past every clip and every locked track: it and all above it are free. */
function laneTop(doc, section) {
    let top = 0;
    for (const p of clipsOf(doc)) if (sectionOf(p) === section) top = Math.max(top, laneOf(p) + 1);
    for (const [k, v] of Object.entries(doc.lanes?.[section] || {})) if (v?.locked && Number.isInteger(Number(k))) top = Math.max(top, Number(k) + 1);
    return top;
}

/** The lowest lane of `section` from `from` up, unlocked and with nothing in [start, end] but what `skip` holds. There always is one. */
function laneFor(doc, section, from, start, end, skip) {
    const top = Math.max(from, laneTop(doc, section));
    let lane = from;
    while (lane < top && !laneFree(doc, section, lane, start, end, skip)) lane += 1;
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
export function placeSpokenLines(doc, { from, lines, voice = null, speed = 1, show, play, typing } = {}) {
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
    // A locked track counts as taken, so a band is never moved onto one.
    const freeAt = (lane) => laneFree(base, 'v', lane, start, end, skip);
    let shiftBy = 0;
    if (bands.length) {
        const fits = (k) => freeAt(wordsFrom + k) && bands.every((b) => freeAt(laneOf(b) + k));
        // Moved far enough, the words and every band are past all clips and locks, so this ends.
        const most = Math.max(0, laneTop(base, 'v') - Math.min(wordsFrom, ...bands.map(laneOf)));
        while (shiftBy < most && !fits(shiftBy)) shiftBy += 1;
    }
    const textLane = bands.length ? wordsFrom + shiftBy : laneFor(base, 'v', wordsFrom, start, end, skip);
    const soundLane = laneFor(base, 'a', 0, start, end, skip);
    const look = lookOf(firstText);
    // A caption's lines stay captions, their typing in step with their voices (S19).
    const caption = firstText.type === 'caption';
    const how = TYPING_TIMINGS.includes(typing) ? typing : (firstText.spoken?.typing === 'first' ? 'first' : 'with');
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
        const words = {
            ...look, id: tid, type: caption ? 'caption' : 'text', name: firstText.name || 'spoken', text: String(l.text).trim(), timing, lane: textLane, group: groups[i],
            enter: i === 0 ? (firstText.enter || still) : still, exit: i === list.length - 1 ? (firstText.exit || still) : still,
            spoken: caption ? { ...spoken, typing: how } : spoken, ...(keepShow ? {} : { visible: false }),
        };
        made.push(caption ? typedLine(words, keepPlay) : words);
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
 * Words typed in Video's Text to speech (D-2026-11-16-01), on the timeline from
 * `at` (the playhead), one line after another in the caption's look, over
 * every picture there is while they run. `lines`: [{ text, seconds, url,
 * name, size }]. With each line's voice made (a url each) they are a spoken
 * layer as Spoken lines makes one, its words and its voice, shown, read or
 * both. Shown only with no voice: the words alone, each for its `seconds`
 * (readingSeconds), marked `shown` so the layer is known to have no voice
 * yet; Read again with a voice gives it one. Answers { doc, layer, ids, start, end }.
 */
export function placeTypedLines(doc, { at = 0, lines, voice = null, speed = 1, show = true, play = true } = {}) {
    const list = (lines || []).filter((l) => l && String(l.text || '').trim()).slice(0, MAX_LINES);
    const start = round(Math.max(0, Number(at) || 0));
    const none = { doc, layer: null, ids: [], start, end: start };
    if (!list.length) return none;
    const voiced = list.every((l) => l.url);
    if (!voiced && play) return none;
    const times = lineTimes(list.map((l) => l.seconds), { at: start });
    const end = times[times.length - 1].end;
    let top = 0;
    for (const p of clipsOf(doc)) if (p.type !== 'sound' && p.timing.end > start + EPS && p.timing.start < end - EPS) top = Math.max(top, laneOf(p) + 1);
    const lane = laneFor(doc, 'v', Math.max(1, top), start, end, new Set());
    const look = captionLook(doc.options?.aspect || 'landscape');
    const still = { kind: 'none', seconds: 0 };
    const taken = { patches: [...doc.patches] };
    if (voiced) {
        // A first line's words where the layer goes, read onto the timeline as Spoken lines reads a title.
        const id = nextId(taken, 'text');
        const seed = { ...look, id, type: 'text', name: 'spoken', text: String(list[0].text).trim(), timing: { start, end: round(Math.max(end, start + MIN_CLIP)) }, lane, enter: still, exit: still };
        return placeSpokenLines({ ...doc, patches: [...doc.patches, seed] }, { from: id, lines: list, voice, speed, show, play });
    }
    const layer = newSpokenLayer(doc);
    const made = list.map((l, i) => {
        const id = nextId(taken, 'text');
        taken.patches.push({ id });
        return { ...look, id, type: 'text', name: 'spoken', text: String(l.text).trim(), timing: times[i], lane, enter: still, exit: still, spoken: { layer, line: i, voice: null, speed, shown: true } };
    });
    return { doc: { ...doc, patches: [...doc.patches, ...made] }, layer, ids: made.map((p) => p.id), start, end };
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
 * as it was read, marked as one spoken layer (two when some were only shown:
 * those have no voice in the take). Their voice is already in the take's
 * sound, so they have no sound of their own. `cues`: [{ start, end, text,
 * show, play }] from the recording's start. Answers { doc, layer, layers, ids }.
 */
export function placeSpokenCues(doc, cues, { at = 0, show = true } = {}) {
    const list = (cues || [])
        .filter((c) => c && String(c.text || '').trim() && Number.isFinite(c.start) && Number.isFinite(c.end) && c.end > c.start)
        .map((c) => ({ start: round(Math.max(0, c.start) + at), end: round(c.end + at), text: String(c.text).trim(), show: c.show !== false, play: c.play !== false }))
        .sort((a, b) => a.start - b.start);
    if (!list.length) return { doc, layer: null, layers: [], ids: [] };
    // One after another they never overlap, as they were read; an end past the next start is cut there.
    for (let i = 0; i + 1 < list.length; i++) if (list[i].end > list[i + 1].start) list[i].end = Math.max(list[i].start + MIN_CLIP, list[i + 1].start);
    // Lines whose voice is in the take and lines that were only shown are layers of their own, so each
    // offers the choices it can take: a voice in the take can be hidden or shown, words alone only shown.
    const used = new Set(clipsOf(doc).map(spokenLayerOf).filter(Boolean));
    const fresh = () => { let n = 1; while (used.has(`sp${n}`)) n += 1; used.add(`sp${n}`); return `sp${n}`; };
    const mixed = list.some((c) => c.play) && list.some((c) => !c.play);
    const voicedLayer = fresh();
    const shownLayer = mixed ? fresh() : voicedLayer;
    const layerOf = (c) => (c.play ? voicedLayer : shownLayer);
    const layer = layerOf(list[0]);
    const lineIn = new Map();
    const start = list[0].start;
    const end = list[list.length - 1].end;
    // Over every picture there is while they run, on a lane of their own: no caption or clip shares it.
    let top = 0;
    for (const p of clipsOf(doc)) if (p.type !== 'sound' && p.timing.end > start + EPS && p.timing.start < end - EPS) top = Math.max(top, laneOf(p) + 1);
    const lane = laneFor(doc, 'v', Math.max(1, top), start, end, new Set());
    const look = captionLook(doc.options?.aspect || 'landscape');
    const taken = { patches: [...doc.patches] };
    const made = list.map((c) => {
        const id = nextId(taken, 'text');
        taken.patches.push({ id });
        const own = layerOf(c);
        const line = lineIn.get(own) || 0;
        lineIn.set(own, line + 1);
        return {
            ...look, id, type: 'text', name: 'spoken', text: c.text, timing: { start: c.start, end: c.end }, lane,
            enter: { kind: 'none', seconds: 0 }, exit: { kind: 'none', seconds: 0 },
            spoken: { layer: own, line, recorded: true, ...(c.play ? {} : { shown: true }) }, ...(show && c.show ? {} : { visible: false }),
        };
    });
    return { doc: { ...doc, patches: [...doc.patches, ...made] }, layer, layers: [...new Set(made.map((p) => p.spoken.layer))], ids: made.map((p) => p.id) };
}

/**
 * The takes a Recorder hands over with the lines it read: the take that
 * carries the recording's sound (`carrier`) counts as sounded, since the
 * lines' voice is in it, even when nothing else was heard at Record.
 */
export function soundedTakes(takes, cues) {
    // Only lines said with their voice into the sound: words shown alone put nothing in it.
    if (!(cues || []).some((c) => c && c.play !== false)) return takes || [];
    return (takes || []).map((k) => (k && k.carrier && !k.hasAudio ? { ...k, hasAudio: true } : k));
}

/**
 * Whether a recording begun now will have spoken lines in its sound: a title
 * with lines to read and a choice that reads them aloud, a voice to read them,
 * and Sound on. Counted as heard when Record is pressed, so the recording keeps
 * its sound track for them.
 */
export function speaksIntoRecording(stack, { voices = false } = {}) {
    if (!voices || !stack || stack.sound !== 'on') return false;
    return (stack.items || []).some((it) => it && it.kind === 'title' && it.spoken?.play !== false && linesOf(it.lines).length > 0);
}

// ── the three choices (D-2026-11-16-01) ─────────────────────────────────────

/**
 * What spoken words do, the same in Video, the Recorder and Live: shown and
 * read aloud, read aloud only, or shown only. Stored as the two switches
 * { show, play } every spoken layer and title already carries.
 */
export const SPEAK_CHOICES = Object.freeze([
    Object.freeze({ id: 'both', show: true, play: true }),
    Object.freeze({ id: 'read', show: false, play: true }),
    Object.freeze({ id: 'show', show: true, play: false }),
]);

/** The choice two switches make: 'both', 'read' or 'show', or null for neither (kept from before, the words silent and unseen). */
export function speakChoiceOf({ show, play } = {}) {
    const s = show !== false;
    const p = play !== false;
    return SPEAK_CHOICES.find((c) => c.show === s && c.play === p)?.id ?? null;
}

/** A choice's two switches, { show, play }; an unknown choice is the first, shown and read. */
export function speakSwitches(id) {
    const c = SPEAK_CHOICES.find((o) => o.id === id) || SPEAK_CHOICES[0];
    return { show: c.show, play: c.play };
}

/** The shortest time a line shown alone stays up, in seconds. */
export const MIN_READING = 1.5;

/**
 * How long a line shown without its voice stays up: about as long as it
 * takes to read, 2.5 words a second at speed 1 (a script written without
 * spaces counted at six characters a word), never less than MIN_READING.
 */
export function readingSeconds(line, { speed = 1 } = {}) {
    const s = String(line ?? '').replace(/\s+/g, ' ').trim();
    if (!s) return 0;
    const words = s.split(' ').length;
    const units = Math.max(words, [...s.replace(/ /g, '')].length / 6);
    const rate = Number(speed) > 0 ? Number(speed) : 1;
    return round(Math.max(MIN_READING, units / 2.5 / rate));
}

/** The most characters a typed line holds before it is said a sentence, then a few words, at a time. */
export const TYPED_LINE_MOST = 80;

/**
 * Typed words as the lines they are said in: each line as typed, but one too
 * long for the band (two rows at most, on a tall stream too) said a sentence
 * at a time, and a sentence still too long a few words at a time. Video's
 * Text to speech lands them the same way.
 */
export function typedLinesOf(text, { most = TYPED_LINE_MOST, language } = {}) {
    const out = [];
    for (const line of linesOf(text)) {
        if (line.length <= most) { out.push(line); continue; }
        for (const sentence of sentencesOfLine(line, language)) {
            if (sentence.length <= most) { out.push(sentence); continue; }
            let row = '';
            for (const word of sentence.split(' ')) {
                if (row && row.length + 1 + word.length > most) { out.push(row); row = word; } else row = row ? `${row} ${word}` : word;
            }
            if (row) out.push(row);
        }
    }
    return out.slice(0, MAX_LINES);
}

/** The studio's typed words, read live: the name its band is kept under, never a source's id (those are "kind-n"). */
export const TYPED_LINE = 'typed';

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

/**
 * The lines a studio reads while recording, kept as cues from the
 * recording's start: [{ start, end, text, show }]. Times are the recording's
 * clock, null while nothing records.
 *
 * - begin(at, text, show, play): a line starts being read; `play` false when its voice is not in the sound.
 * - end(at): it finished; kept if the recording was running.
 * - take(end): the recording stopped at `end`; a line still being read is
 *   kept to there, and every cue is handed over and forgotten.
 * - clear(): a recording begins; nothing from before is kept, and a line
 *   being read now counts from the recording's start.
 */
export function spokenCueLog() {
    let cues = [];
    let open = null;
    const keep = (o, end) => {
        const start = o.at ?? 0;
        if (end > start) cues.push({ start, end, text: o.text, show: o.show, play: o.play });
    };
    return {
        begin(at, text, show = true, play = true) {
            open = { at: Number.isFinite(at) ? at : null, text, show, play, closed: false };
        },
        end(at) {
            const o = open;
            open = null;
            if (o && !o.closed && Number.isFinite(at)) keep(o, at);
        },
        take(end = null) {
            if (open && !open.closed && Number.isFinite(end)) {
                keep(open, end);
                open.closed = true;
            }
            const out = cues;
            cues = [];
            return out;
        },
        clear() {
            cues = [];
            if (open) open = { ...open, at: null, closed: false };
        },
    };
}
