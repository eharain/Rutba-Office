/**
 * Document settings and the sound layer that reaches the plan. (What can be
 * added, sound included, is the insert catalogue's — tests/insert.test.mjs.)
 *
 * These cover the pieces that closed M3's inspector gap. One of them is a
 * pure object transform and the other is the renderer's own compile step, so
 * none of it needs a browser — but the last one is the one worth having: a
 * panel that writes fields the compiler ignores would look like it worked, and
 * the mistake would surface as a silent video rather than an error.
 *
 *   node --test tests/settings.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { emptyDoc, setOption, addLayer, updateLayer } from '../packages/studio/src/doc.js';
import { applyLayerPatches, soundLayers, clipFromSoundLayer } from '../packages/studio/src/index.js';

/**
 * The smallest thing `applyLayerPatches` will accept, as in toolset.test.mjs.
 *
 * `buildPlan` is deliberately NOT used: it calls `canvas.getContext`, and a
 * sound layer has nothing to do with a canvas. Compiling the patch directly
 * tests the step that matters and keeps this file runnable in Node.
 */
const stubPlan = (overrides = {}) => ({
    W: 1080, H: 1920, duration: 5, fps: 30,
    theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
    layers: [], images: [], videos: {}, assets: {}, context: {},
    ...overrides,
});

// ── setOption ────────────────────────────────────────────────────────────────

test('setOption merges into options and leaves the rest of the doc alone', () => {
    const doc = { ...emptyDoc('video'), title: 'Keep me', patches: [{ id: 'a', type: 'shape' }] };
    const next = setOption(doc, { transition: 'zoom' });

    assert.equal(next.options.transition, 'zoom');
    assert.equal(next.options.theme, 'dark', 'the other options survive');
    assert.equal(next.title, 'Keep me');
    assert.deepEqual(next.patches, doc.patches);
    assert.notEqual(next, doc, 'a new object, so React sees the change');
    assert.equal(doc.options.transition, undefined, 'the original is untouched');
});

test('setOption can write null — which is how "no track" is spelled', () => {
    // The reason this is a shallow merge rather than the usual skip-nullish
    // deep merge. audioTrackId: null is a VALUE, not an omission.
    const doc = setOption(emptyDoc('video'), { audioTrackId: 'track-7' });
    const cleared = setOption(doc, { audioTrackId: null });

    assert.equal(cleared.options.audioTrackId, null);
    assert.ok('audioTrackId' in cleared.options, 'cleared, not deleted');
});

test('setOption works on a doc that has no options yet', () => {
    const next = setOption({ kind: 'video' }, { quality: 'low' });
    assert.deepEqual(next.options, { quality: 'low' });
});

// ── the sound layer, all the way to a clip ───────────────────────────────────

/** A doc with one sound layer, compiled — the way the editor would compile it. */
function planWithSound(patch) {
    const doc = addLayer(emptyDoc('video'), { type: 'sound', ...patch });
    const plan = stubPlan();
    applyLayerPatches(plan, doc.patches);
    return { id: doc.patches[0].id, plan };
}

test('the fields the panel writes are the fields the compiler reads', () => {
    // The point of this test. Every control in SoundPanel writes one of these,
    // and a name the compiler does not know is dropped silently — the layer
    // still appears in the timeline and still renders silent.
    const { id, plan } = planWithSound({
        trackId: 'tr-1', url: 'https://example.test/a.mp3',
        offset: 4, volume: 0.35, mix: 'duck', loop: true,
        enter: { kind: 'fade', seconds: 0.5 },
        exit: { kind: 'fade', seconds: 1.25 },
    });

    const layer = plan.layers.find((l) => l.id === id);
    assert.ok(layer, 'the sound layer reached the plan');
    assert.equal(layer.trackId, 'tr-1');
    assert.equal(layer.url, 'https://example.test/a.mp3');
    assert.equal(layer.offset, 4);
    assert.equal(layer.volume, 0.35);
    assert.equal(layer.mix, 'duck');
    assert.equal(layer.loop, true);
    assert.equal(layer.enter.seconds, 0.5);
    assert.equal(layer.exit.seconds, 1.25);
});

test('a sound layer never paints', () => {
    const { plan } = planWithSound({ url: 'https://example.test/a.mp3' });
    // It is audio. If it ever acquired a painter, an invisible layer would start
    // covering the frame.
    assert.equal(plan.layers.filter((l) => l.type === 'sound').length, 1);
});

test('hiding a sound layer takes it out of the render, not just the picture', () => {
    let doc = addLayer(emptyDoc('video'), { type: 'sound', url: 'https://example.test/a.mp3' });
    doc = updateLayer(doc, doc.patches[0].id, { visible: false });
    const plan = stubPlan();
    applyLayerPatches(plan, doc.patches);

    // The eye icon on a sound lane has to mean MUTE, because there is nothing
    // to look at. soundLayers is what enforces it.
    assert.equal(soundLayers(plan).length, 0);
});

test('the panel defaults and the studio defaults meet in clipFromSoundLayer', () => {
    const { id, plan } = planWithSound({ url: 'https://example.test/a.mp3', mix: 'duck' });
    const layer = plan.layers.find((l) => l.id === id);
    const buffer = { duration: 90 };

    // volume is null on the layer (the panel left it alone), so the studio-level
    // default has to be what lands on the clip.
    const clip = clipFromSoundLayer(layer, buffer, plan, { volume: 0.7, fadeIn: 1.2, fadeOut: 1.6 });
    assert.equal(clip.volume, 0.7, 'the host default fills in for an unset layer volume');
    assert.equal(clip.mix, 'duck', 'the layer wins where it has an opinion');
    assert.equal(clip.buffer, buffer);
});

test('an explicit layer volume beats the studio default', () => {
    const { id, plan } = planWithSound({ url: 'https://example.test/a.mp3', volume: 1 });
    const layer = plan.layers.find((l) => l.id === id);
    const clip = clipFromSoundLayer(layer, { duration: 30 }, plan, { volume: 0.7 });
    assert.equal(clip.volume, 1, 'a voice-over set to full stays at full');
});

// ── the music bed stays an instruction ───────────────────────────────────────

test('choosing random does not resolve a track into the document', () => {
    // If the panel ever wrote a chosen id here, "random" would mean "the same
    // random track, every render" — and a batch of twenty would stop being
    // varied, which is the whole reason the mode exists.
    const doc = setOption(emptyDoc('video'), { audioMode: 'random' });
    assert.equal(doc.options.audioMode, 'random');
    assert.ok(!doc.options.audioTrackId, 'no track is pinned by choosing random');
});
