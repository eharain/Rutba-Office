/**
 * The image editor's render path.
 *
 * AN IMAGE IS A ONE-FRAME VIDEO. `paintFrame(ctx, plan, t)` already paints any
 * instant of a composition onto a canvas — layers, fractional geometry,
 * keyframes, `{token}` binding, brand marks and all. What a video adds on top
 * is `MediaRecorder`, real-time pacing and audio, and an image needs none of
 * it. So the image editor is `buildPlan` → `paintFrame` → encode one bitmap,
 * and this file is the small amount of glue that says so.
 *
 * That is why M2 is smaller than the plan assumed. The ERP had no image editor
 * at all, but it had a painter general enough to be one, and reusing it means
 * an image creative and a video creative are the same document with the same
 * layer semantics — a template can serve both, and a caption that reads well in
 * one reads the same in the other.
 *
 * WHAT AN IMAGE PROJECT DOES DIFFERENTLY:
 *  - It has no duration, so `at` picks which instant to freeze (default 0).
 *  - It has no audio; sound layers are ignored rather than an error, so a
 *    template shared with video degrades instead of refusing.
 *  - It encodes with `canvas.toBlob`, which is not real time — a batch of
 *    images is bounded by pixels, not by seconds of video.
 */

import { ASPECTS, buildPlan, paintFrame } from './index.js';

/** Per-platform still sizes. The three video aspects plus the shapes only stills use. */
export const IMAGE_SIZES = {
    ...Object.fromEntries(Object.entries(ASPECTS).map(([k, v]) => [k, { ...v }])),
    portrait: { key: 'portrait', label: 'Portrait 4:5', hint: 'Instagram feed (max height)', width: 1080, height: 1350 },
    wide: { key: 'wide', label: 'Wide 1.91:1', hint: 'Link previews · OG images', width: 1200, height: 628 },
    story: { key: 'story', label: 'Story 9:16', hint: 'Stories · Status', width: 1080, height: 1920 },
};

const MIME = { png: 'image/png', jpeg: 'image/jpeg', webp: 'image/webp' };

/**
 * Which still formats this engine will actually produce.
 *
 * `toBlob` silently falls back to PNG for a type it does not support rather
 * than failing, so asking for WebP and getting a PNG named `.webp` is a real
 * outcome. Probed once, the same way the video path probes MediaRecorder.
 */
export function supportedImageFormats(canvasFactory = () => document.createElement('canvas')) {
    const c = canvasFactory();
    c.width = 1; c.height = 1;
    return Object.keys(MIME).filter((fmt) => {
        try { return c.toDataURL(MIME[fmt]).startsWith(`data:${MIME[fmt]}`); } catch { return false; }
    });
}

/**
 * Build a plan for a still.
 *
 * `size` may name an IMAGE_SIZES key or give `{width, height}` directly. The
 * renderer works in aspects, so a custom size is passed through as one — every
 * layer's geometry is fractional, which is exactly what makes that safe.
 */
export function buildImagePlan({ canvas, size = 'square', images = [], title, body, logo, options, layerPatches, context, assets }) {
    const named = typeof size === 'string' ? IMAGE_SIZES[size] : null;
    const dims = named ?? (size && size.width && size.height ? size : IMAGE_SIZES.square);

    const args = {
        canvas,
        images,
        title,
        body,
        logo,
        layerPatches,
        context,
        assets,
        options: {
            ...options,
            // A still has no motion, so anything that only reads at a moving
            // playhead is off: the progress bar, the edge dip, Ken Burns.
            showProgress: false,
            edgeFadeSeconds: 0,
            kenBurns: false,
        },
    };

    // A shape the video renderer already knows needs nothing special.
    if (named && ASPECTS[named.key]) {
        return buildPlan({ ...args, options: { ...args.options, aspect: named.key } });
    }

    return buildPlanAtSize(args, dims);
}

/**
 * Compile at a frame size `ASPECTS` does not contain.
 *
 * THIS CANNOT BE DONE BY RESIZING AFTERWARDS. Stored geometry is fractional,
 * but `compileLayers` RESOLVES those fractions to pixels while it builds the
 * plan — the painters take compile-time pixels, which is what keeps them fast
 * and simple. Setting `plan.W`/`plan.H` after the fact would therefore leave
 * every layer, the caption band and the photo stage measured for the previous
 * frame, and the result looks subtly wrong rather than obviously broken: text
 * off-centre, a logo drifting out of its corner.
 *
 * So the size has to be known BEFORE the plan is built, and `buildPlan` reads
 * it from `ASPECTS[opts.aspect]`. The honest way to add one is to put it there,
 * build, and put the table back. That mutation is safe because `buildPlan` is
 * synchronous: nothing else can observe the table between these two lines, in a
 * single-threaded runtime. `finally` covers the case where it throws.
 */
function buildPlanAtSize(args, dims) {
    const key = `__image_${dims.width}x${dims.height}`;
    const previous = ASPECTS[key];
    ASPECTS[key] = { key, label: `${dims.width}×${dims.height}`, hint: 'still', width: dims.width, height: dims.height };
    try {
        return buildPlan({ ...args, options: { ...args.options, aspect: key } });
    } finally {
        if (previous === undefined) delete ASPECTS[key]; else ASPECTS[key] = previous;
    }
}

/**
 * Paint one frame and encode it.
 *
 * `at` is the instant to freeze. It defaults to 0 rather than to the middle,
 * because a still built from an image template has nothing happening on its
 * clock, and a still built from a VIDEO template should show that template's
 * opening — which is the frame a person picturing it has in mind.
 *
 * Returns `{ blob, extension, mimeType, width, height }`, matching what
 * `renderVideo` returns, so a caller can upload either without branching.
 */
export async function renderImage({ canvas, plan, at = 0, format = 'png', quality = 0.92 }) {
    const mimeType = MIME[format] ?? MIME.png;

    // Re-assert the plan's size on the canvas before painting.
    //
    // `buildImagePlan` sizes the canvas, but a caller holding several plans —
    // which `renderImageSet` does, and which is the natural way to offer "this
    // creative at four sizes" — will have had it resized by whichever plan was
    // built last. Painting a 1080x1350 plan onto a canvas some other plan left
    // at 640x480 produces a wrong-sized image and NO error at all: the layers
    // are compiled in pixels, so they simply land off the edge.
    //
    // Cheap, idempotent, and it makes a plan self-sufficient: given a plan, the
    // canvas it is handed no longer has to be the one it was built with.
    if (canvas.width !== plan.W || canvas.height !== plan.H) {
        canvas.width = plan.W;
        canvas.height = plan.H;
    }

    const ctx = canvas.getContext('2d');
    paintFrame(ctx, plan, Math.max(0, Math.min(at, plan.duration ?? 0)));

    const blob = await new Promise((resolve, reject) => {
        canvas.toBlob(
            (b) => (b ? resolve(b) : reject(new Error(`could not encode ${format}`))),
            mimeType,
            format === 'png' ? undefined : quality,
        );
    });

    // Report what was ACTUALLY produced. toBlob falls back to PNG rather than
    // failing on an unsupported type, so trusting the request would hand the
    // caller a PNG to store under a .webp name.
    const actual = blob.type || mimeType;
    return {
        blob,
        mimeType: actual,
        extension: (Object.entries(MIME).find(([, m]) => m === actual)?.[0] ?? 'png').replace('jpeg', 'jpg'),
        width: canvas.width,
        height: canvas.height,
    };
}

/**
 * One composition at several sizes — the batch a social post actually needs.
 *
 * Each size gets its own plan rather than one plan scaled, because layout is
 * not scale: a chip clear of the caption at 9:16 can land under it at 1:1, and
 * `lib/aspects.js` exists to report exactly that. Compiling per size is what
 * makes the check meaningful.
 */
export async function renderImageSet({ canvas, sizes, format = 'png', quality = 0.92, at = 0, ...args }) {
    const out = [];
    for (const size of sizes) {
        const plan = buildImagePlan({ canvas, size, ...args });
        const rendered = await renderImage({ canvas, plan, at, format, quality });
        out.push({ size: typeof size === 'string' ? size : `${rendered.width}x${rendered.height}`, ...rendered });
    }
    return out;
}
