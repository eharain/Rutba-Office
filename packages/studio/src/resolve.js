/**
 * How a STORED document's references become the loaded media a plan needs.
 *
 * Two hosts render a creative — the editor's browser and the render worker's
 * headless page — and both start from the same JSON document: a list of
 * picture urls, patches whose layers name clips and sounds by url or by
 * library trackId, and a music bed that lives in `options`. Each host used to
 * derive "what must be fetched, and what clips does renderVideo get" on its
 * own, and they drifted exactly the way two implementations of one rule
 * always do: the worker rendered every video silent and pictureless, and the
 * editor once rendered every library-picked voice-over silent (the 2026-08-26
 * audit). This module is the one copy of that rule.
 *
 * WHAT IS DELIBERATELY NOT HERE: the fetching itself. A page must go through
 * a same-origin proxy or the canvas taints; a worker holds a service
 * credential and asks Node; and only the worker may query the database for a
 * track row. So every I/O arrives injected — `loadTrack`, `lookupTrack`,
 * `randomTrack` — and this module stays a pure description of WHAT to
 * resolve, testable in Node with fakes (tests/resolve.test.mjs).
 *
 * THE RANDOM BED RESOLVES FRESH ON EVERY CALL, and nothing is ever written
 * back into the document. Pinning the drawn track's id would turn
 * `audioMode: 'random'` into "the same random track, every time" — and the
 * varied batch is the whole point (PLAN.md M3). The host's `randomTrack`
 * decides the population; both hosts must scope it `org OR global`, never
 * every track on the instance.
 */

import { soundLayers, clipFromSoundLayer } from './index.js';

/**
 * The urls a stored document needs fetched and decoded before it can compile.
 *
 *   imageUrls  the photo lane (`doc.images`) — ORDER MATTERS and duplicates
 *              stay: photo layers reference these by index, so this list must
 *              arrive at `buildPlan({ images })` in exactly this order
 *   clipUrls   video-clip layers — deduped, because two uses of one take
 *              should share one decode pipeline (`buildPlan({ videos })` is
 *              keyed by url)
 *   assetUrls  image layers that carry their own url (a swapped logo, an
 *              imported picture) — deduped, keyed by url into
 *              `buildPlan({ assets })`
 *
 * Sound urls are not listed here: audio is not needed to COMPILE a plan, and
 * resolving it needs the plan's own clock — see resolveAudioClips.
 */
export function mediaManifest(doc) {
    const patches = Array.isArray(doc?.patches) ? doc.patches : [];
    const uniq = (urls) => [...new Set(urls)];
    return {
        imageUrls: (Array.isArray(doc?.images) ? doc.images : [])
            .filter((u) => typeof u === 'string' && u),
        clipUrls: uniq(patches.filter((p) => p?.type === 'video' && p.url).map((p) => p.url)),
        assetUrls: uniq(patches.filter((p) => p?.type === 'image' && p.url).map((p) => p.url)),
    };
}

/**
 * The playable url of a library track row. A track is EITHER an uploaded file
 * (`media_ref`) or a foreign url (`source_url`); the upload wins when both
 * exist, because it is the copy this instance controls.
 */
export function trackUrl(track) {
    return track?.media_ref || track?.source_url || null;
}

/**
 * Everything a document wants heard, as the `audio` argument renderVideo
 * takes — or null when there is nothing to hear.
 *
 * Failures are swallowed clip by clip, deliberately: a track that will not
 * load loses the music, not the video. A silent render is a recoverable
 * disappointment; a failed one is the full real-time length of the video,
 * spent for nothing.
 *
 * `io`:
 *   loadTrack(url)    → AudioBuffer (the host's transport + decoder)
 *   lookupTrack(id)   → a library track row, org-scoped by the host
 *   randomTrack()     → one ACTIVE track row at random, scoped org OR global
 *   random()          → [0,1) — injectable so the random-start offset is
 *                       testable; defaults to Math.random
 */
export async function resolveAudioClips(plan, options, io) {
    const opts = options || {};
    const { loadTrack, lookupTrack, randomTrack, random = Math.random } = io || {};
    const clips = [];

    // Sound layers first — they carry their own timing. A layer names its
    // sound in one of TWO ways — a url, or a library trackId — and reading
    // only the url was the bug that rendered every library-picked voice-over
    // silent. A layer is not resolved until both spellings have been tried.
    for (const layer of soundLayers(plan)) {
        let url = layer.url || null;
        if (!url && layer.trackId && lookupTrack) {
            try { url = trackUrl(await lookupTrack(layer.trackId)); } catch { url = null; }
        }
        if (!url) continue;
        try {
            const buffer = await loadTrack(url);
            clips.push(clipFromSoundLayer(layer, buffer, plan, {
                volume: opts.audioVolume, fadeIn: opts.audioFadeIn, fadeOut: opts.audioFadeOut,
            }));
        } catch { /* this clip is silent; the rest still plays */ }
    }

    // Then the bed, which lives in options rather than as a layer. 'random'
    // asks the host rather than picking here, so the editor and the render
    // worker draw from the SAME scoped population — and nothing is resolved
    // into the document, so a batch stays varied.
    if (opts.audioMode && opts.audioMode !== 'none') {
        try {
            const track = opts.audioMode === 'random'
                ? await randomTrack?.()
                : (opts.audioTrackId && lookupTrack ? await lookupTrack(opts.audioTrackId) : null);
            const url = trackUrl(track);
            if (url) {
                const buffer = await loadTrack(url);
                clips.push({
                    buffer,
                    start: 0,
                    end: plan.duration,
                    offset: opts.audioRandomStart && buffer.duration > plan.duration
                        ? random() * (buffer.duration - plan.duration)
                        : (track.start_offset || 0),
                    volume: track.volume ?? opts.audioVolume ?? 0.7,
                    fadeIn: opts.audioFadeIn ?? 1.2,
                    fadeOut: opts.audioFadeOut ?? 1.6,
                    loop: true,
                });
            }
        } catch { /* no bed; the video still renders */ }
    }

    return clips.length ? { clips } : null;
}
