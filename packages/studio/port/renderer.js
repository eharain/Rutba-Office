/**
 * The renderer, wired to this app's media transport — and to Studio's toolset.
 *
 * TWO THINGS HAPPEN ON IMPORT, and both are load-bearing.
 *
 * 1. The media transport is installed. `@rutba-studio/editor-core` ships with
 *    NO transport on purpose: it serves two hosts that fetch bytes in genuinely
 *    different ways (a page must go through a same-origin proxy or the canvas
 *    is tainted; the render worker holds a service credential and reads the
 *    media namespace directly), and that seam is the only reason one file can
 *    serve both. This is the page's half.
 *
 * 2. The toolset is registered. wordart / shapes / adjust register themselves
 *    on import, and paintFrame SILENTLY SKIPS a layer type it has no painter
 *    for — a deliberate forward-compatibility rule, so a recipe from a newer
 *    editor degrades instead of exploding. The failure mode when you forget
 *    this import is therefore not an error: it is a headline that shows in the
 *    preview and is missing from the render. Every page reaches the renderer
 *    through THIS module so that cannot happen.
 *
 * Everything app-specific stays on this side of the line.
 */
import { configureMediaFetch, loadImages as loadImageSet } from "../../../packages/editor-core/index.js";

/**
 * The surface, named explicitly rather than star-exported.
 *
 * This started as three `export *` lines and did not survive contact with a
 * bundler: `toolset.js` re-exports `registerLayerType` and
 * `registeredLayerTypes` from `index.js`, so two of the stars offered the same
 * two names. An ambiguous star export is dropped rather than resolved, and what
 * came out the other side was a module missing `wordArtResize` with no clue as
 * to why. Naming everything costs a list and removes the whole class of
 * problem — and the list doubles as the answer to "what may a page use?".
 */
export {
    // shapes of the output
    ASPECTS, THEMES, DEFAULTS,
    // compile and paint
    buildPlan, applyLayerPatches, compileLayers, paintFrame,
    // media
    loadImage, releaseImages, loadVideo, releaseVideos,
    loadAudioTrack, audioContext, isAudioFile, isVideoFile,
    configureMediaFetch,
    // editor primitives
    layerBounds, hitTestLayers, layerHandles, hitTestHandles, scaleFromDrag, resizePatch,
    withLayerStateAt,
    // video output
    renderVideo, startAudioPreview, soundLayers, clipFromSoundLayer,
    pickMimeType, extensionFor, unsupportedReason, videoFileName,
    // the extension registry
    registerLayerType, registeredLayerTypes,
} from "../../../packages/editor-core/index.js";

/**
 * Load pictures and answer with THE ENTRIES ALONE — an array, not the
 * `{ images, failures }` report editor-core's loadImages returns.
 *
 * THE BUG THIS SEAM EXISTS TO END: six call sites across this app passed that
 * report object straight into `buildPlan({ images })`, which indexes an array
 * — so `images.length` was undefined, the photo count went NaN, and NOT ONE
 * PICTURE EVER DREW, in the editor, the share page, the deck thumbnails, the
 * presenter or the exports. No error anywhere: a photo layer whose bitmap is
 * missing simply doesn't paint, by design. Two authors made the identical
 * mistake, which is the tell that the trap belonged to the seam rather than
 * to either of them — so the app-side wrapper now returns what every caller
 * here actually wants, and the next call site written the obvious way is
 * right by construction.
 *
 * The failure list is dropped on purpose: this app's posture everywhere is
 * that a missing picture loses the picture, not the render. A caller that
 * genuinely wants the report imports editor-core directly, as the render
 * worker's page does.
 */
export async function loadImages(urls, opts) {
    const { images } = await loadImageSet(urls, opts);
    return images;
}

/**
 * The still path. Re-exported HERE rather than imported directly by the pages
 * that use it, for the same reason everything else is: importing
 * `editor-core/image.js` on its own gets you a renderer with no toolset
 * registered, and the symptom is a headline that previews and does not export.
 */
export { buildImagePlan, renderImage, renderImageSet, supportedImageFormats, IMAGE_SIZES }
    from "../../../packages/editor-core/image.js";

/**
 * Stored-document resolution — the one copy of "what must be fetched, and
 * what does renderVideo get to hear", shared with the render worker's page so
 * the two hosts cannot drift (see the module's own header).
 */
export { mediaManifest, resolveAudioClips, trackUrl }
    from "../../../packages/editor-core/resolve.js";

/**
 * The toolset. Importing this module is what REGISTERS word art, shapes and
 * adjustment layers into the renderer — see the note at the top of the file.
 */
export {
    WORDART, SHAPE, ADJUST,
    WORDART_PRESETS, WARPS, GEOMETRIES, ADJUST_PRESETS,
    ADD_MENU, applyAdjustment, wordArtResize,
    SHAPE_GROUPS, GEOMETRY_LABELS, paintShapeTile,
    ICONS, ICON_NAMES, ICON_GROUPS, ICON_LABELS,
    LIST_KINDS, LIST_LABELS,
    CHART_KINDS, CHART_LABELS, CHART_PALETTE,
} from "../../../packages/editor-core/toolset.js";

// The proxy refuses hosts it has not been told serve our media. Music tracks
// are the exception — a library of FOREIGN urls is the point of the feature —
// so the proxy verifies those against the audio library instead, and needs the
// caller's identity to read it. Pages set this once they have a session.
let _auth = null;

/** Give the proxy the caller's identity, so it can check foreign track urls. */
export function setMediaAuth({ jwt, appRole } = {}) {
    _auth = jwt ? { jwt, appRole: appRole || null } : null;
}

function authHeaders() {
    const headers = {};
    if (_auth?.jwt) {
        headers.Authorization = `Bearer ${_auth.jwt}`;
        headers["X-Rutba-App"] = "studio";
        if (_auth.appRole) headers["X-Rutba-App-Role"] = _auth.appRole;
    }
    return headers;
}

/**
 * Fetch media as a blob, through this app's origin.
 *
 * Returning a BLOB rather than a URL is the contract the renderer asks for, and
 * the reason is the taint: a blob: URL minted from these bytes is same-origin,
 * so the canvas that draws it stays clean and both captureStream() (for video)
 * and getImageData() (for every adjustment layer) keep working.
 */
async function fetchMediaViaProxy(url, { signal, range } = {}) {
    const headers = authHeaders();
    // Only `bytes=<n>-<n>` gets past the proxy's own matcher, so it is the only
    // shape built here. An open-ended range is legal and means "to the end".
    if (range) headers.Range = `bytes=${range.start}-${range.end ?? ""}`;
    const res = await fetch(`/api/media-proxy?url=${encodeURIComponent(url)}`, { signal, headers });
    if (!res.ok) {
        let detail = `HTTP ${res.status}`;
        try { detail = (await res.json())?.error || detail; } catch { /* keep the status */ }
        throw new Error(detail);
    }
    return res.blob();
}

/**
 * How large a track is, and whether it can be fetched in part. One HEAD round
 * trip, against downloading an hour of audio to play thirty seconds of it.
 *
 * NOT handed to the renderer. The pinned editor-core predates
 * `configureMediaProbe` (the ERP's copy grew it after the fork point, and
 * re-pinning is a deliberate act, not a sync — see harness/BASELINE.md), so
 * windowing is the CALLER'S business here. The audio picker uses it directly.
 */
export async function probeMediaViaProxy(url, { signal } = {}) {
    const res = await fetch(`/api/media-proxy?url=${encodeURIComponent(url)}`, {
        method: "HEAD", signal, headers: authHeaders(),
    });
    // A probe that cannot answer is not an error — it means "no windowing", and
    // the caller falls back to the whole file on its own terms.
    if (!res.ok) return { size: 0, acceptsRanges: false };
    const size = Number(res.headers.get("content-length"));
    return {
        size: Number.isFinite(size) ? size : 0,
        acceptsRanges: /bytes/i.test(res.headers.get("accept-ranges") || ""),
    };
}

configureMediaFetch(fetchMediaViaProxy);

export { fetchMediaViaProxy };
