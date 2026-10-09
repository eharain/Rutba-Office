/**
 * Stored-document resolution — packages/editor-core/resolve.js.
 *
 * ONE implementation decides what a document needs fetched and what
 * renderVideo gets to hear, and BOTH hosts run it: the editor's export bar
 * and the render worker's page. These tests drive it with fakes for the
 * injected I/O, which is the point of the injection — every rule here is
 * host-independent, and a host that disagreed with a rule would be the drift
 * this module exists to end.
 *
 * Two rules earn their tests especially:
 *
 *   - a layer names its sound by url OR by library trackId, and both
 *     spellings must be tried — reading only the url is the audit bug that
 *     rendered every library-picked voice-over silent;
 *   - `audioMode: 'random'` resolves FRESH on every call and writes nothing
 *     back — pinning the drawn track would turn a twenty-render batch into
 *     twenty copies of one clip, and the varied batch is M3's whole point.
 *
 *   node --test tests/resolve.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { mediaManifest, resolveAudioClips, trackUrl } from '../packages/studio/src/resolve.js';

// ── the manifest ─────────────────────────────────────────────────────────────

test('the photo lane keeps its order and its duplicates', () => {
    // Photo layers reference doc.images BY INDEX. Deduping or sorting here
    // would silently re-point every photo at the wrong picture.
    const doc = { images: ['b', 'a', 'b'], patches: [] };
    assert.deepEqual(mediaManifest(doc).imageUrls, ['b', 'a', 'b']);
});

test('clip and asset urls dedupe; sound urls are not the manifest\'s business', () => {
    const doc = {
        images: [],
        patches: [
            { type: 'video', url: 'v1' }, { type: 'video', url: 'v1' },
            { type: 'image', url: 'i1' }, { type: 'image', url: 'i1' },
            { type: 'image' },                     // the compiled logo — no url
            { type: 'sound', url: 's1' },          // audio resolves later, on the plan
        ],
    };
    const m = mediaManifest(doc);
    assert.deepEqual(m.clipUrls, ['v1'], 'two uses of one take share one decode');
    assert.deepEqual(m.assetUrls, ['i1']);
    assert.ok(!JSON.stringify(m).includes('s1'));
});

test('a malformed document answers empty lists rather than throwing', () => {
    for (const doc of [null, {}, { images: 'nope', patches: 'nope' }]) {
        assert.deepEqual(mediaManifest(doc), { imageUrls: [], clipUrls: [], assetUrls: [] });
    }
});

// ── the track url ────────────────────────────────────────────────────────────

test('an uploaded file beats a foreign url — it is the copy this instance controls', () => {
    assert.equal(trackUrl({ media_ref: 'ours', source_url: 'theirs' }), 'ours');
    assert.equal(trackUrl({ source_url: 'theirs' }), 'theirs');
    assert.equal(trackUrl({}), null);
    assert.equal(trackUrl(null), null);
});

// ── sound layers ─────────────────────────────────────────────────────────────

const soundLayer = (extra = {}) => ({ type: 'sound', visible: true, ...extra });
const planWith = (...layers) => ({ duration: 12, layers });
const buf = (duration) => ({ duration });

test('a layer with a url needs no lookup', async () => {
    const audio = await resolveAudioClips(planWith(soundLayer({ url: 'https://m.test/a.wav' })), {}, {
        loadTrack: async (url) => { assert.equal(url, 'https://m.test/a.wav'); return buf(3); },
        lookupTrack: async () => { throw new Error('must not be called'); },
    });
    assert.equal(audio.clips.length, 1);
});

test('a layer with only a trackId resolves through the library — the audit bug, pinned', async () => {
    const audio = await resolveAudioClips(planWith(soundLayer({ trackId: 't1' })), {}, {
        loadTrack: async (url) => { assert.equal(url, 'https://m.test/t1.mp3'); return buf(3); },
        lookupTrack: async (id) => (id === 't1' ? { media_ref: 'https://m.test/t1.mp3' } : null),
    });
    assert.equal(audio.clips.length, 1);
});

test('a foreign track resolves to its source url', async () => {
    const audio = await resolveAudioClips(planWith(soundLayer({ trackId: 't2' })), {}, {
        loadTrack: async () => buf(3),
        lookupTrack: async () => ({ source_url: 'https://elsewhere.test/t2.mp3' }),
    });
    assert.equal(audio.clips.length, 1);
});

test('a failed lookup or decode is a silent clip, not a failed render', async () => {
    // Three layers: one lookup throws, one decode throws, one lands. The video
    // still renders with the one that worked — a silent clip is a recoverable
    // disappointment, a failed render is the video's full length for nothing.
    const plan = planWith(
        soundLayer({ trackId: 'down' }),
        soundLayer({ url: 'https://m.test/bad.wav' }),
        soundLayer({ url: 'https://m.test/good.wav' }),
    );
    const audio = await resolveAudioClips(plan, {}, {
        loadTrack: async (url) => { if (/bad/.test(url)) throw new Error('no decode'); return buf(3); },
        lookupTrack: async () => { throw new Error('library down'); },
    });
    assert.equal(audio.clips.length, 1);
});

test('a layer with neither spelling, and a hidden layer, resolve to nothing', async () => {
    const plan = planWith(soundLayer({}), soundLayer({ url: 'u', visible: false }));
    const audio = await resolveAudioClips(plan, {}, {
        loadTrack: async () => { throw new Error('must not be called'); },
    });
    assert.equal(audio, null, 'nothing to hear answers null, not { clips: [] }');
});

test('a clip carries its layer\'s own clock and mix', async () => {
    const layer = soundLayer({
        url: 'u', timing: { start: 2, end: 9 }, offset: 1.5, mix: 'duck', volume: 0.4,
    });
    const audio = await resolveAudioClips(planWith(layer), {}, { loadTrack: async () => buf(30) });
    const clip = audio.clips[0];
    assert.equal(clip.start, 2);
    assert.equal(clip.end, 9);
    assert.equal(clip.offset, 1.5);
    assert.equal(clip.mix, 'duck');
    assert.equal(clip.volume, 0.4);
});

// ── the music bed ────────────────────────────────────────────────────────────

test('no mode, or mode none, asks for no bed', async () => {
    for (const options of [{}, { audioMode: 'none' }]) {
        const audio = await resolveAudioClips(planWith(), options, {
            loadTrack: async () => { throw new Error('must not be called'); },
            randomTrack: async () => { throw new Error('must not be called'); },
        });
        assert.equal(audio, null);
    }
});

test('a picked bed loops under the whole video with the track\'s own settings', async () => {
    const options = { audioMode: 'pick', audioTrackId: 't9' };
    const audio = await resolveAudioClips(planWith(), options, {
        loadTrack: async () => buf(200),
        lookupTrack: async (id) => (id === 't9'
            ? { media_ref: 'u', volume: 0.5, start_offset: 12.5 }
            : null),
    });
    const clip = audio.clips[0];
    assert.equal(clip.start, 0);
    assert.equal(clip.end, 12, 'the bed ends with the video');
    assert.equal(clip.loop, true);
    assert.equal(clip.volume, 0.5, 'the track\'s stored volume wins');
    assert.equal(clip.offset, 12.5, 'a start point someone chose beats a computed one');
});

test('bed volume falls back track → options → 0.7', async () => {
    const io = (track) => ({ loadTrack: async () => buf(20), lookupTrack: async () => track });
    const opts = { audioMode: 'pick', audioTrackId: 't' };

    let audio = await resolveAudioClips(planWith(), { ...opts, audioVolume: 0.3 }, io({ media_ref: 'u' }));
    assert.equal(audio.clips[0].volume, 0.3);

    audio = await resolveAudioClips(planWith(), opts, io({ media_ref: 'u' }));
    assert.equal(audio.clips[0].volume, 0.7);
});

test('a random bed draws FRESH each call and writes nothing back', async () => {
    // This is the varied batch: twenty renders of one doc must be twenty
    // draws. Pinning the first draw into the document — the tempting "fix" —
    // is exactly what this asserts cannot happen.
    const options = Object.freeze({ audioMode: 'random' });
    let draws = 0;
    const io = {
        loadTrack: async () => buf(20),
        lookupTrack: async () => { throw new Error('random must not resolve by id'); },
        randomTrack: async () => ({ media_ref: `u${draws += 1}` }),
    };
    await resolveAudioClips(planWith(), options, io);
    await resolveAudioClips(planWith(), options, io);
    assert.equal(draws, 2, 'one draw per render');
    assert.deepEqual(options, { audioMode: 'random' }, 'the document is untouched');
});

test('an empty library loses the bed, not the render', async () => {
    const audio = await resolveAudioClips(planWith(), { audioMode: 'random' }, {
        loadTrack: async () => { throw new Error('must not be called'); },
        randomTrack: async () => null,
    });
    assert.equal(audio, null);
});

test('a random start point is drawn from the slack, or falls back to the track\'s own', async () => {
    const options = { audioMode: 'pick', audioTrackId: 't', audioRandomStart: true };
    const io = (duration) => ({
        loadTrack: async () => buf(duration),
        lookupTrack: async () => ({ media_ref: 'u', start_offset: 4 }),
        random: () => 0.5,
    });

    // 30s of track under a 12s video: 18s of slack, drawn at 0.5 → 9.
    let audio = await resolveAudioClips(planWith(), options, io(30));
    assert.equal(audio.clips[0].offset, 9);

    // A track shorter than the video has no slack to draw from.
    audio = await resolveAudioClips(planWith(), options, io(8));
    assert.equal(audio.clips[0].offset, 4);
});
