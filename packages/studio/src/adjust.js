/**
 * Adjustment layers — the per-pixel half of the image editor.
 *
 * The renderer already has `layer.filter`, which compiles to a CSS filter
 * string. That covers brightness, contrast, saturate, grayscale, sepia, hue and
 * blur, it is free (the compositor does it), and it should stay the first
 * choice for anything it can express. What it CANNOT express is most of what
 * "advanced image editing" means: levels with a black point, a tone curve, a
 * white balance, vibrance that protects skin, a vignette, a sharpen.
 *
 * So this is the other tool. An `adjust` layer changes THE PIXELS ALREADY
 * PAINTED BENEATH IT — the classic adjustment layer — which falls out of the
 * renderer's architecture for free: paintFrame draws layers in z order onto one
 * context, so "everything below" is just "the canvas, right now".
 *
 * HOW IT GETS OPACITY, MASKS AND BLEND MODES FOR FREE. putImageData ignores
 * globalAlpha, the clip region, the transform and the composite operation — it
 * is a raw write, and an adjustment layer built on it could not be faded or
 * masked at all. So the adjusted pixels go to an offscreen canvas and come back
 * through drawImage, which honours all four. That single indirection is why an
 * adjustment can be 40% strong, masked to an ellipse, and set to `multiply`,
 * with none of that written here.
 *
 * COST, HONESTLY. This reads and writes every pixel in its region. On a still
 * that is nothing. On a 1080x1920 video at 30fps it is roughly 62 million pixel
 * ops per second per full-frame adjustment, and it will not hold real time —
 * and this renderer records in real time, so a dropped frame is a dropped frame
 * in the file. Bound the region (`fw`/`fh` default to the whole frame precisely
 * so that choice is visible), or keep adjustments to image projects. The editor
 * is expected to warn rather than to silently produce a stuttering video.
 */

import { registerLayerType } from './index.js';

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const clamp255 = (v) => (v < 0 ? 0 : v > 255 ? 255 : v);
const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);

/** Named starting points, so the panel opens on a look rather than on zeros. */
export const ADJUST_PRESETS = {
    none: { label: 'None', adjust: {} },
    punch: { label: 'Punch', adjust: { contrast: 0.22, vibrance: 0.35, sharpen: 0.4 } },
    matte: { label: 'Matte', adjust: { levels: { outBlack: 22, outWhite: 240 }, saturation: -0.12 } },
    warm: { label: 'Warm', adjust: { temperature: 0.28, vibrance: 0.2 } },
    cool: { label: 'Cool', adjust: { temperature: -0.25, tint: 0.06 } },
    mono: { label: 'Mono', adjust: { saturation: -1, contrast: 0.16 } },
    fade: { label: 'Fade', adjust: { levels: { outBlack: 34, gamma: 1.12 }, saturation: -0.25, vignette: 0.18 } },
    noir: { label: 'Noir', adjust: { saturation: -1, contrast: 0.42, vignette: 0.4 } },
};

// ── the lookup table ─────────────────────────────────────────────────────────

/**
 * Every adjustment that maps a channel value to a channel value, collapsed into
 * three 256-entry tables.
 *
 * This is the whole reason a full-frame adjustment is viable at all: levels,
 * gamma, curves, exposure, brightness, contrast and white balance are ALL point
 * operations, so composing them per pixel would redo the same arithmetic two
 * million times for 256 distinct answers. Built once per frame, then the pixel
 * loop is three array reads.
 *
 * Order matters and follows a darkroom rather than a whim: levels set the black
 * and white points first (everything downstream assumes a normalised range),
 * the curve reshapes the tones inside it, exposure and contrast scale the
 * result, and white balance is last because it is a per-channel gain on
 * finished tone.
 */
function buildLuts(a) {
    const lutR = new Uint8ClampedArray(256);
    const lutG = new Uint8ClampedArray(256);
    const lutB = new Uint8ClampedArray(256);

    const lv = a.levels || {};
    const inBlack = clamp(num(lv.inBlack, 0), 0, 254);
    const inWhite = clamp(num(lv.inWhite, 255), inBlack + 1, 255);
    const gamma = clamp(num(lv.gamma, 1), 0.1, 9.99);
    const outBlack = clamp(num(lv.outBlack, 0), 0, 255);
    const outWhite = clamp(num(lv.outWhite, 255), 0, 255);

    const exposure = num(a.exposure, 0);          // stops
    const brightness = num(a.brightness, 0);      // -1 .. 1
    const contrast = num(a.contrast, 0);          // -1 .. 1
    const gain = Math.pow(2, exposure);
    // The standard contrast slope, so ±1 is a strong but not destroyed image.
    const slope = (1.0 + clamp(contrast, -0.99, 0.99)) / (1.0 - clamp(contrast, -0.99, 0.99));

    const curveR = curveLut(a.curves && (a.curves.r || a.curves.rgb));
    const curveG = curveLut(a.curves && (a.curves.g || a.curves.rgb));
    const curveB = curveLut(a.curves && (a.curves.b || a.curves.rgb));

    // White balance as per-channel gain. Temperature moves red against blue,
    // tint moves green against magenta — the two axes a camera exposes.
    const temp = clamp(num(a.temperature, 0), -1, 1);
    const tint = clamp(num(a.tint, 0), -1, 1);
    const gR = 1 + temp * 0.32;
    const gB = 1 - temp * 0.32;
    const gG = 1 + tint * 0.24;

    for (let i = 0; i < 256; i++) {
        // levels in
        let v = (i - inBlack) / (inWhite - inBlack);
        v = v < 0 ? 0 : v > 1 ? 1 : v;
        if (gamma !== 1) v = Math.pow(v, 1 / gamma);
        // levels out
        v = (outBlack + v * (outWhite - outBlack)) / 255;
        // exposure + brightness + contrast, about mid grey
        v = v * gain + brightness;
        v = (v - 0.5) * slope + 0.5;
        const base = clamp255(Math.round(v * 255));
        lutR[i] = clamp255(Math.round(curveR[base] * gR));
        lutG[i] = clamp255(Math.round(curveG[base] * gG));
        lutB[i] = clamp255(Math.round(curveB[base] * gB));
    }

    return { lutR, lutG, lutB };
}

/**
 * Control points → a 256-entry curve, by monotone linear interpolation.
 *
 * Points are `[[inputA, outputA], ...]` in 0..255. Linear rather than spline on
 * purpose: a Catmull-Rom through user points can overshoot past 255 and back,
 * which shows up as a bright rim on a smooth gradient, and a curve tool that
 * invents contrast the user did not ask for is worse than a slightly stiff one.
 */
function curveLut(points) {
    const out = new Uint8ClampedArray(256);
    if (!Array.isArray(points) || points.length < 2) {
        for (let i = 0; i < 256; i++) out[i] = i;
        return out;
    }
    const pts = points
        .map((p) => [clamp(num(p[0], 0), 0, 255), clamp(num(p[1], 0), 0, 255)])
        .sort((p, q) => p[0] - q[0]);
    let seg = 0;
    for (let i = 0; i < 256; i++) {
        while (seg < pts.length - 2 && i > pts[seg + 1][0]) seg++;
        const [x0, y0] = pts[seg];
        const [x1, y1] = pts[seg + 1];
        if (i <= x0) { out[i] = y0; continue; }
        if (i >= x1) { out[i] = y1; continue; }
        out[i] = Math.round(y0 + ((i - x0) / (x1 - x0)) * (y1 - y0));
    }
    return out;
}

// ── the pixel passes ─────────────────────────────────────────────────────────

/**
 * Saturation and vibrance. Not point operations — they need all three channels
 * — so they cannot join the LUT and get their own pass.
 *
 * Vibrance differs from saturation in the one way that matters in practice: it
 * scales its effect DOWN as a pixel is already saturated, so a blue sky
 * intensifies while skin, which is already fairly saturated, is left roughly
 * alone. That is the whole reason to have both sliders.
 */
function applyColour(data, saturation, vibrance) {
    const sat = clamp(saturation, -1, 4);
    const vib = clamp(vibrance, -1, 2);
    for (let i = 0; i < data.length; i += 4) {
        const r = data[i]; const g = data[i + 1]; const b = data[i + 2];
        // Rec. 709 luma — the renderer paints for screens, not for print.
        const l = 0.2126 * r + 0.7152 * g + 0.0722 * b;
        let k = 1 + sat;
        if (vib !== 0) {
            const mx = r > g ? (r > b ? r : b) : (g > b ? g : b);
            const mn = r < g ? (r < b ? r : b) : (g < b ? g : b);
            const current = mx === 0 ? 0 : (mx - mn) / mx;   // how saturated already
            k += vib * (1 - current);
        }
        data[i] = clamp255(l + (r - l) * k);
        data[i + 1] = clamp255(l + (g - l) * k);
        data[i + 2] = clamp255(l + (b - l) * k);
    }
}

/**
 * A 3x3 sharpen, strength-blended against the original.
 *
 * Blended rather than applied at a fixed kernel because an unsharp mask at
 * full strength haloes every edge; the strength slider has to mean something
 * between "off" and "obviously over-sharpened".
 */
function applySharpen(src, w, h, strength) {
    const k = clamp(strength, 0, 1);
    if (k <= 0) return src;
    const out = new Uint8ClampedArray(src.length);
    out.set(src);
    const centre = 1 + 4 * k;
    const side = -k;
    for (let y = 1; y < h - 1; y++) {
        for (let x = 1; x < w - 1; x++) {
            const i = (y * w + x) * 4;
            for (let c = 0; c < 3; c++) {
                const p = i + c;
                out[p] = clamp255(
                    src[p] * centre
                    + src[p - 4] * side + src[p + 4] * side
                    + src[p - w * 4] * side + src[p + w * 4] * side,
                );
            }
        }
    }
    return out;
}

/** A radial darkening toward the corners. Cheap, and the reason `fade` reads as film. */
function applyVignette(data, w, h, amount) {
    const k = clamp(amount, -1, 1);
    if (k === 0) return;
    const cx = w / 2; const cy = h / 2;
    const maxD = Math.hypot(cx, cy);
    for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
            const d = Math.hypot(x - cx, y - cy) / maxD;
            // Squared falloff keeps the centre clean and rolls off near the edge.
            const f = 1 - k * d * d;
            const i = (y * w + x) * 4;
            data[i] = clamp255(data[i] * f);
            data[i + 1] = clamp255(data[i + 1] * f);
            data[i + 2] = clamp255(data[i + 2] * f);
        }
    }
}

/**
 * Apply an adjustment to raw RGBA in place (except sharpen, which needs a
 * second buffer and is therefore returned).
 *
 * Exported because it is useful without a layer at all: the image editor's
 * thumbnail strip and the "before/after" toggle both want to adjust a bitmap
 * they already hold, and neither of them is painting a frame.
 */
export function applyAdjustment(data, w, h, adjust) {
    const a = adjust || {};
    let out = data;

    const needsLut = a.levels || a.curves || a.exposure || a.brightness || a.contrast
        || a.temperature || a.tint;
    if (needsLut) {
        const { lutR, lutG, lutB } = buildLuts(a);
        for (let i = 0; i < out.length; i += 4) {
            out[i] = lutR[out[i]];
            out[i + 1] = lutG[out[i + 1]];
            out[i + 2] = lutB[out[i + 2]];
        }
    }

    if (a.saturation || a.vibrance) applyColour(out, num(a.saturation, 0), num(a.vibrance, 0));
    if (a.invert) for (let i = 0; i < out.length; i += 4) {
        out[i] = 255 - out[i]; out[i + 1] = 255 - out[i + 1]; out[i + 2] = 255 - out[i + 2];
    }
    if (a.posterize) {
        const steps = Math.max(2, Math.round(num(a.posterize, 8)));
        const q = 255 / (steps - 1);
        for (let i = 0; i < out.length; i += 4) {
            out[i] = Math.round(out[i] / q) * q;
            out[i + 1] = Math.round(out[i + 1] / q) * q;
            out[i + 2] = Math.round(out[i + 2] / q) * q;
        }
    }
    if (a.vignette) applyVignette(out, w, h, num(a.vignette, 0));
    if (a.sharpen) out = applySharpen(out, w, h, num(a.sharpen, 0));

    return out;
}

// ── the layer ────────────────────────────────────────────────────────────────

/**
 * The offscreen surface the adjusted pixels travel back through.
 *
 * Cached on the layer and resized in place: allocating a canvas the size of the
 * frame on every video frame is a garbage-collection pause you can see in the
 * recording, and the recording is real time — a pause is a lost frame.
 */
function surfaceFor(layer, w, h) {
    let s = layer._adjSurface;
    if (!s) {
        s = typeof OffscreenCanvas === 'function'
            ? new OffscreenCanvas(w, h)
            : document.createElement('canvas');
        layer._adjSurface = s;
    }
    if (s.width !== w || s.height !== h) { s.width = w; s.height = h; }
    return s;
}

function paintAdjust(ctx, plan, layer) {
    const region = adjustBounds(ctx, plan, layer);
    if (!region) return;

    // Snap to whole pixels: getImageData takes integers, and a fractional box
    // that rounds differently on read and write leaves a one-pixel seam.
    const x = Math.max(0, Math.floor(region.x));
    const y = Math.max(0, Math.floor(region.y));
    const w = Math.min(plan.W - x, Math.ceil(region.w));
    const h = Math.min(plan.H - y, Math.ceil(region.h));
    if (w <= 0 || h <= 0) return;

    let img;
    try {
        img = ctx.getImageData(x, y, w, h);
    } catch {
        // A tainted canvas. The renderer goes to great lengths to avoid one
        // (every image arrives as a blob through fetchMedia for exactly this
        // reason), but a host that skipped that seam should lose the
        // adjustment, not the whole render.
        return;
    }

    const adjusted = applyAdjustment(img.data, w, h, layer.adjust);
    // Copy back into the ImageData we were HANDED rather than constructing a
    // new one. Sharpen is the only pass that needs a second buffer, so this is
    // usually the same array and the set() is a no-op — and it means this file
    // never names the `ImageData` constructor, which is not guaranteed to exist
    // in every context this renderer is expected to run in (the render worker's
    // is a headless browser today, but the seam exists so it need not stay one).
    if (adjusted !== img.data) img.data.set(adjusted);
    const surface = surfaceFor(layer, w, h);
    const sctx = surface.getContext('2d');
    sctx.putImageData(img, 0, 0);

    // Soft edges live HERE rather than in the renderer's clipToMask, which is
    // hard-edged by construction: an alpha ramp needs a surface to multiply
    // into, and this is the only place in the pipeline that has one.
    const feather = layer.mask && num(layer.mask.feather, 0);
    if (feather > 0) {
        sctx.save();
        sctx.globalCompositeOperation = 'destination-in';
        const r = Math.max(w, h) / 2;
        const g = sctx.createRadialGradient(w / 2, h / 2, r * (1 - clamp(feather, 0, 1)), w / 2, h / 2, r);
        g.addColorStop(0, 'rgba(0,0,0,1)');
        g.addColorStop(1, 'rgba(0,0,0,0)');
        sctx.fillStyle = g;
        sctx.fillRect(0, 0, w, h);
        sctx.restore();
    }

    // Back through drawImage, which — unlike putImageData — respects the
    // globalAlpha, clip and composite operation the frame wrapper set up.
    ctx.drawImage(surface, x, y, w, h);
}

/** The region. Absent fractions mean the whole frame, which is the common case. */
function adjustBounds(ctx, plan, layer) {
    if (layer.fw == null && layer.fh == null) return { x: 0, y: 0, w: plan.W, h: plan.H };
    const w = plan.W * (layer.fw == null ? 1 : layer.fw);
    const h = layer.fh == null ? w : plan.H * layer.fh;
    return { x: layer.x - w / 2, y: layer.y - h / 2, w, h };
}

function compileAdjust(patch, cx) {
    const { W, H } = cx;
    const preset = ADJUST_PRESETS[patch.preset];
    const env = cx.envelope(patch.anim);
    return {
        name: patch.name || 'Adjustment',
        adjust: { ...(preset ? preset.adjust : {}), ...(patch.adjust || {}) },
        preset: patch.preset || null,
        visible: patch.visible !== false,
        timing: patch.timing || null,
        enter: patch.enter || env.enter,
        exit: patch.exit || env.exit,
        z: patch.z == null ? cx.nextZ() : patch.z,
        x: W * (patch.fx == null ? 0.5 : patch.fx),
        y: H * (patch.fy == null ? 0.5 : patch.fy),
        fx: patch.fx == null ? 0.5 : patch.fx,
        fy: patch.fy == null ? 0.5 : patch.fy,
        fw: patch.fw == null ? null : patch.fw,
        fh: patch.fh == null ? null : patch.fh,
        opacity: patch.opacity == null ? 1 : patch.opacity,
        blend: patch.blend || null,
        mask: patch.mask || null,
        // An adjustment layer never rotates: it reads an axis-aligned rectangle
        // of the canvas, and a rotated read is a different, much slower thing.
        rot: 0,
        anim: patch.anim || 'none',
        keys: cx.keys(patch.keys),
    };
}

export const ADJUST = registerLayerType('adjust', {
    paint: paintAdjust,
    compile: compileAdjust,
    bounds: adjustBounds,
});
