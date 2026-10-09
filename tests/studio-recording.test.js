/**
 * What a recorded take becomes once it is on the creative.
 *
 * The capture itself belongs to `@rutba/ui`'s RecorderDialog and is exercised
 * where it lives — it needs a microphone and a MediaRecorder, neither of which
 * a test has. What IS testable, and where the decisions actually are, is the
 * shape of what gets placed afterwards: how many layers a clip becomes, whether
 * they line up, and whether the renderer can see them.
 *
 * The one that earns its keep is the last group. A `video` patch compiles with
 * `visible: patch.visible !== false && !!entry` — no decoded entry in
 * `plan.videos`, no layer — so before the editor learned to load clips, a
 * recorded take landed as a lane that drew nothing and said nothing about why.
 *
 *   node --test tests/recording.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { applyLayerPatches, soundLayers } from '../packages/studio/src/index.js';
import { emptyDoc, addLayer } from '../packages/studio/src/doc.js';

const stubPlan = (over = {}) => ({
    W: 1080, H: 1920, duration: 20, fps: 30,
    theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
    layers: [], images: [], videos: {}, assets: {}, context: {}, ...over,
});

const compile = (patches, over) => {
    const plan = stubPlan(over);
    applyLayerPatches(plan, patches);
    return plan;
};

const REF = 'https://media.test/take-1.webm';

/**
 * The placement RecordBar performs, as data — the same two shapes it appends.
 * Kept here rather than imported because the component is JSX and this is about
 * the patches, not the buttons.
 */
function placed(mode, { start = 0, seconds = 4 } = {}) {
    const end = start + seconds;
    if (mode === 'video') {
        return [
            { type: 'video', url: REF, name: 'take', timing: { start, end } },
            { type: 'sound', url: REF, name: 'take (sound)', mix: 'duck', volume: 1, timing: { start, end } },
        ];
    }
    return [{ type: 'sound', url: REF, name: 'take', mix: 'duck', volume: 1, timing: { start, end } }];
}

// ── an audio take ───────────────────────────────────────────────────────────

test('an audio take is one sound layer, pointing at its own url', () => {
    const plan = compile(placed('audio').map((p, i) => ({ ...p, id: `s${i}` })));
    const sounds = soundLayers(plan);
    assert.equal(sounds.length, 1);
    assert.equal(sounds[0].url, REF, 'the layer carries the url, not a library id');
});

test('a recorded voice ducks the bed by default', () => {
    // The default IS the feature, as with the Insert palette's voice-over: a
    // voice competing with music at equal level is the commonest way a take
    // comes out unusable, and `mix` is the field nobody would go looking for.
    const plan = compile(placed('audio').map((p, i) => ({ ...p, id: `s${i}` })));
    assert.equal(soundLayers(plan)[0].mix, 'duck');
});

test('a take lands at the playhead, not at zero', () => {
    const plan = compile(placed('audio', { start: 6.5, seconds: 3 }).map((p, i) => ({ ...p, id: `s${i}` })));
    const l = soundLayers(plan)[0];
    assert.equal(l.timing.start, 6.5, 'somebody who scrubbed to 6.5s meant "here"');
    assert.equal(l.timing.end, 9.5);
});

// ── a video take ────────────────────────────────────────────────────────────

test('a clip is TWO layers on one url — picture and sound', () => {
    // The renderer's own split. It is what lets the picture and the sound be
    // trimmed independently, and it makes muting a take a matter of deleting
    // one lane rather than hunting for a checkbox.
    const patches = placed('video').map((p, i) => ({ ...p, id: `v${i}` }));
    const plan = compile(patches, { videos: { [REF]: { el: {}, duration: 4 } } });

    const picture = plan.layers.filter((l) => l.type === 'video');
    const sound = plan.layers.filter((l) => l.type === 'sound');
    assert.equal(picture.length, 1);
    assert.equal(sound.length, 1);
    assert.equal(picture[0].url, sound[0].url, 'both halves name the same file');
});

test('both halves of a clip start and end together', () => {
    // They can be trimmed apart afterwards; they must not ARRIVE apart, or the
    // first thing anyone hears is sound over the wrong picture.
    const patches = placed('video', { start: 2, seconds: 5 }).map((p, i) => ({ ...p, id: `v${i}` }));
    const plan = compile(patches, { videos: { [REF]: { el: {}, duration: 5 } } });
    const [pic] = plan.layers.filter((l) => l.type === 'video');
    const [snd] = plan.layers.filter((l) => l.type === 'sound');
    assert.deepEqual(pic.timing, snd.timing);
});

// ── the gap that made this necessary ────────────────────────────────────────

test('a clip with no decoded entry compiles INVISIBLE', () => {
    // Pinning the trap rather than the fix. This is why useCreative has to load
    // videos into plan.videos: without an entry the layer is silently not drawn,
    // and there is no error anywhere to explain the empty frame.
    const patches = placed('video').map((p, i) => ({ ...p, id: `v${i}` }));
    const plan = compile(patches, { videos: {} });
    const [pic] = plan.layers.filter((l) => l.type === 'video');
    assert.equal(pic.visible, false, 'no entry, no picture — and no complaint');
});

test('the same clip is visible once its entry is decoded', () => {
    const patches = placed('video').map((p, i) => ({ ...p, id: `v${i}` }));
    const plan = compile(patches, { videos: { [REF]: { el: {}, duration: 4 } } });
    const [pic] = plan.layers.filter((l) => l.type === 'video');
    assert.equal(pic.visible, true);
});

test('the sound half needs no decoded entry — it is resolved by the host', () => {
    // Which is why an audio take worked before clips did: sound never consults
    // plan.videos, the export bar loads it through loadAudioTrack instead.
    const plan = compile(placed('video').map((p, i) => ({ ...p, id: `v${i}` })), { videos: {} });
    assert.equal(soundLayers(plan).length, 1);
});

// ── the url the loader keys on ──────────────────────────────────────────────

test('two clips on one file decode once, not twice', () => {
    // useCreative keys its loader on the DEDUPED url set. A creative that uses
    // the same take twice must not open two <video> elements for it — each one
    // holds a decode pipeline, not just a buffer.
    let doc = emptyDoc('video');
    doc = addLayer(doc, { type: 'video', url: REF, timing: { start: 0, end: 3 } });
    doc = addLayer(doc, { type: 'video', url: REF, timing: { start: 5, end: 8 } });
    doc = addLayer(doc, { type: 'video', url: 'https://media.test/other.webm' });

    const urls = [...new Set(doc.patches.filter((p) => p.type === 'video' && p.url).map((p) => p.url))];
    assert.equal(urls.length, 2);
});

test('a clip layer with no url is not asked for', () => {
    // An unfinished layer must not send the loader after `undefined`.
    let doc = emptyDoc('video');
    doc = addLayer(doc, { type: 'video' });
    const urls = doc.patches.filter((p) => p.type === 'video' && p.url).map((p) => p.url);
    assert.deepEqual(urls, []);
});

// Resolving a sound layer to its bytes — the trackId-or-url rule the audit's
// bug taught — moved into editor-core's shared resolver so the editor and the
// render worker run ONE copy of it. Pinned in tests/resolve.test.mjs.
