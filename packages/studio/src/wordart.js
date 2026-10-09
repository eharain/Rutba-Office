/**
 * Word art — text as a graphic object.
 *
 * WHY THIS IS ITS OWN LAYER TYPE. The renderer already has a `text` layer, and
 * it is deliberately plain: one line, one fill, one font, painted exactly the
 * way the ERP's footer painted it. That identity is load-bearing — the
 * extraction gate compares it byte for byte against the frozen baseline — so it
 * is the one thing in this package that must not grow features.
 *
 * Word art is the other half of that trade. Studio is a MEDIA TOOL: it has no
 * product catalogue to bind a headline to, so the headline has to be worth
 * making by hand. That means gradients, outlines, extrudes and shape warps, and
 * none of it belongs anywhere near the frozen painter. A second type keeps both
 * promises at once.
 *
 * WHAT IT PAINTS
 *   fill      solid | linear gradient | radial gradient, across the art's own
 *             box, so the ramp follows the words rather than the frame
 *   strokes   any number, painted widest-first UNDER the fill, which is how an
 *             outline reads as an outline instead of eating the letterforms
 *   shadow    drop | glow | long (a fading trail) | 3d (a solid extrude)
 *   warp      arc, wave, rise, bulge — per character, with each character
 *             ROTATED to the tangent, because a curve of upright letters looks
 *             like a mistake and a curve of tangent letters looks like a logo
 *
 * ZERO DEPENDENCIES, like the rest of this package. It paints with canvas 2D
 * and nothing else, so the render worker runs it unchanged.
 *
 * COST. Layout is the expensive half (per-character measurement, then the warp
 * transform) and a video repaints 30 times a second, so layout is memoised on
 * the layer under a signature of everything that can change it — the same trick
 * paintQr uses for its matrix. A static headline measures once per render.
 */

import { registerLayerType } from './index.js';

/** The warps a word-art layer can take. `none` is a plain baseline. */
export const WARPS = ['none', 'arc', 'wave', 'rise', 'bulge'];

/** Ready-made looks, so the first click produces something rather than nothing. */
export const PRESETS = {
    plain: { label: 'Plain', style: {} },
    outline: {
        label: 'Outline',
        style: { fill: { kind: 'solid', color: '#ffffff' }, strokes: [{ width: 0.07, color: '#000000' }] },
    },
    sticker: {
        label: 'Sticker',
        style: {
            fill: { kind: 'solid', color: '#ffffff' },
            strokes: [{ width: 0.16, color: '#000000' }, { width: 0.09, color: '#ffc107' }],
            shadow: { kind: 'drop', blur: 0.06, dx: 0.02, dy: 0.03, color: 'rgba(0,0,0,0.45)' },
        },
    },
    gold: {
        label: 'Gold',
        style: {
            fill: { kind: 'linear', angle: 90, stops: [[0, '#fff3b0'], [0.45, '#e0a800'], [0.55, '#8a5a00'], [1, '#ffd76e']] },
            strokes: [{ width: 0.05, color: '#4a3000' }],
            shadow: { kind: 'drop', blur: 0.05, dx: 0, dy: 0.025, color: 'rgba(0,0,0,0.5)' },
        },
    },
    neon: {
        label: 'Neon',
        style: {
            fill: { kind: 'solid', color: '#ffffff' },
            strokes: [{ width: 0.035, color: '#22d3ee' }],
            shadow: { kind: 'glow', blur: 0.14, color: '#22d3ee' },
        },
    },
    extrude: {
        label: '3D',
        style: {
            fill: { kind: 'solid', color: '#ffc107' },
            shadow: { kind: '3d', depth: 0.09, angle: 45, color: '#7a4b00' },
        },
    },
    arch: {
        label: 'Arch',
        style: { warp: 'arc', curve: 0.55, fill: { kind: 'solid', color: '#ffffff' }, strokes: [{ width: 0.06, color: '#000000' }] },
    },
};

const DEFAULT_STYLE = {
    family: '"Segoe UI", system-ui, Roboto, "Helvetica Neue", Arial, sans-serif',
    weight: 800,
    italic: false,
    uppercase: false,
    // Fractions of the font size, so a look survives being resized.
    tracking: 0,      // letter spacing
    leading: 1.12,    // line height
    align: 'center',
    fill: { kind: 'solid', color: '#ffffff' },
    strokes: [],
    shadow: null,
    warp: 'none',
    curve: 0.5,       // arc: -1 (deep valley) .. 1 (deep arch); wave: amplitude
    frequency: 1.5,   // wave only: cycles across the line
};

/** The widest arc one line may bend into — beyond this it overlaps itself. */
const MAX_ARC = Math.PI * 0.92;

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);

function styleOf(layer) {
    const s = layer.style || {};
    return {
        ...DEFAULT_STYLE,
        ...s,
        fill: s.fill || DEFAULT_STYLE.fill,
        strokes: Array.isArray(s.strokes) ? s.strokes : [],
    };
}

function fontOf(st, sizePx) {
    return (st.italic ? 'italic ' : '') + (st.weight || 800) + ' ' + Math.round(sizePx) + 'px ' + st.family;
}

// ── layout ───────────────────────────────────────────────────────────────────

/**
 * Measure the art and place every character.
 *
 * Returns `{ lines, box }` in a LOCAL space whose origin is the anchor point,
 * so the painter translates once and draws. Each placed character carries its
 * own transform, because a warp rotates and scales characters independently —
 * there is no single transform that bends a string.
 */
function layout(ctx, layer) {
    const st = styleOf(layer);
    const sizePx = layer.sizePx;
    const font = fontOf(st, sizePx);
    const raw = String(layer.text == null ? '' : layer.text);
    const source = st.uppercase ? raw.toUpperCase() : raw;
    const rows = source.split('\n');
    const track = num(st.tracking, 0) * sizePx;
    const lineStep = num(st.leading, 1.12) * sizePx;

    ctx.save();
    ctx.font = font;

    const lines = rows.map((rowText, rowIndex) => {
        const chars = [...rowText];
        // Advance per character INCLUDING tracking. Measured one at a time
        // rather than as a whole run: the warp needs each character's own
        // centre, and a kerned measurement of the run cannot give one.
        const advances = chars.map((ch) => ctx.measureText(ch).width + track);
        const width = advances.reduce((a, b) => a + b, 0) - (chars.length ? track : 0);
        return { chars, advances, width, y: rowIndex * lineStep };
    });

    ctx.restore();

    const widest = lines.reduce((m, l) => Math.max(m, l.width), 0) || 1;

    for (const line of lines) {
        const alignShift = st.align === 'left' ? 0
            : st.align === 'right' ? (widest - line.width)
            : (widest - line.width) / 2;
        let d = 0;                                   // distance along the line
        line.placed = line.chars.map((ch, i) => {
            const adv = line.advances[i];
            const centre = d + (adv - track) / 2;    // centre of this glyph
            d += adv;
            const base = { ch, x: alignShift + centre - widest / 2, y: line.y, rot: 0, scale: 1 };
            return warpChar(base, { st, line, widest, sizePx });
        });
    }

    // The bounding box, taken from the PLACED characters rather than the flat
    // metrics: a warp moves things outside the box they started in.
    let minX = Infinity; let maxX = -Infinity; let minY = Infinity; let maxY = -Infinity;
    for (const line of lines) {
        for (const p of line.placed) {
            const half = (sizePx * 0.62) * p.scale;   // generous: caps + descenders
            minX = Math.min(minX, p.x - half); maxX = Math.max(maxX, p.x + half);
            minY = Math.min(minY, p.y - sizePx * 0.82 * p.scale);
            maxY = Math.max(maxY, p.y + sizePx * 0.30 * p.scale);
        }
    }
    if (!Number.isFinite(minX)) { minX = 0; maxX = 0; minY = 0; maxY = 0; }

    return {
        lines,
        box: { x: minX, y: minY, w: Math.max(1, maxX - minX), h: Math.max(1, maxY - minY) },
        font, sizePx, st,
    };
}

/**
 * Move one character according to the layer's warp.
 *
 * `p` is the flat placement: x measured from the line's centre, y the line's
 * baseline. Every warp is a displacement of that, so 'none' is genuinely the
 * identity and a warp can be switched off without a second code path.
 */
function warpChar(p, { st, line, widest, sizePx }) {
    const kind = st.warp;
    if (!kind || kind === 'none') return p;
    // Position along the line, roughly -0.5 (start) .. 0.5 (end).
    const u = line.width > 0 ? p.x / Math.max(1, widest) : 0;

    if (kind === 'arc') {
        const curve = clamp(num(st.curve, 0.5), -1, 1);
        const theta = Math.abs(curve) * MAX_ARC;
        if (theta < 1e-3) return p;
        const dir = curve >= 0 ? 1 : -1;
        const r = line.width / theta;
        // a = 0 at the line's centre; the character sits on the arc at a.
        const a = (p.x / Math.max(1, line.width)) * theta;
        return {
            ...p,
            x: r * Math.sin(a),
            // dir = +1 puts the ENDS lower than the middle — a rainbow. That
            // sign is the whole difference between an arch and a valley.
            y: p.y + dir * r * (1 - Math.cos(a)),
            rot: dir * a,
        };
    }

    if (kind === 'wave') {
        const amp = clamp(num(st.curve, 0.5), -1, 1) * sizePx * 0.5;
        const freq = Math.max(0.2, num(st.frequency, 1.5));
        const phase = u * Math.PI * 2 * freq;
        const slope = Math.cos(phase) * amp * ((Math.PI * 2 * freq) / Math.max(1, widest));
        return { ...p, y: p.y + Math.sin(phase) * amp, rot: Math.atan2(slope, 1) };
    }

    if (kind === 'rise') {
        const lift = clamp(num(st.curve, 0.5), -1, 1) * sizePx;
        return { ...p, y: p.y + (u + 0.5) * -lift, rot: Math.atan2(-lift, Math.max(1, widest)) };
    }

    if (kind === 'bulge') {
        // Middle characters grow and the ends shrink (reversed for a negative
        // curve). Scale ONLY — a bulge that also moved characters would open
        // gaps the eye reads as broken tracking.
        const amount = clamp(num(st.curve, 0.5), -1, 1);
        return { ...p, scale: 1 + amount * 0.45 * Math.cos(u * Math.PI) };
    }

    return p;
}

/** Memoised layout — recomputed only when something that changes it changes. */
function layoutFor(ctx, layer) {
    const st = styleOf(layer);
    const sig = JSON.stringify([layer.text, layer.sizePx, st.family, st.weight, st.italic,
        st.uppercase, st.tracking, st.leading, st.align, st.warp, st.curve, st.frequency]);
    if (layer._wa && layer._wa.sig === sig) return layer._wa.value;
    const value = layout(ctx, layer);
    layer._wa = { sig, value };
    return value;
}

// ── painting ─────────────────────────────────────────────────────────────────

/** A fill spec → something assignable to ctx.fillStyle, in the art's own box. */
function fillStyle(ctx, fill, box) {
    if (!fill || !fill.kind || fill.kind === 'solid') return (fill && fill.color) || '#ffffff';
    const stops = Array.isArray(fill.stops) && fill.stops.length
        ? fill.stops
        : [[0, '#ffffff'], [1, '#888888']];
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    let g;
    if (fill.kind === 'radial') {
        const r = Math.max(box.w, box.h) / 2;
        g = ctx.createRadialGradient(cx, cy, r * 0.05, cx, cy, r);
    } else {
        // Angle in degrees; 0 = left-to-right, measured clockwise.
        const a = (num(fill.angle, 90) * Math.PI) / 180;
        const half = (Math.abs(Math.cos(a)) * box.w + Math.abs(Math.sin(a)) * box.h) / 2;
        g = ctx.createLinearGradient(
            cx - Math.cos(a) * half, cy - Math.sin(a) * half,
            cx + Math.cos(a) * half, cy + Math.sin(a) * half,
        );
    }
    for (const [at, color] of stops) g.addColorStop(clamp(num(at, 0), 0, 1), color);
    return g;
}

/** Run `draw` once per character, with that character's transform applied. */
function forEachChar(ctx, laid, draw) {
    for (const line of laid.lines) {
        for (const p of line.placed) {
            ctx.save();
            ctx.translate(p.x, p.y);
            if (p.rot) ctx.rotate(p.rot);
            if (p.scale !== 1) ctx.scale(p.scale, p.scale);
            draw(p.ch);
            ctx.restore();
        }
    }
}

function paintWordArt(ctx, plan, layer) {
    if (!layer.text) return;
    const laid = layoutFor(ctx, layer);
    const st = laid.st;
    const sizePx = laid.sizePx;

    ctx.save();
    ctx.translate(layer.x, layer.y);
    ctx.font = laid.font;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'alphabetic';
    ctx.lineJoin = 'round';
    ctx.miterLimit = 2;

    const sh = st.shadow;

    // 1. The extrude, if any. Painted FIRST and as its own solid body, because
    //    depth is geometry rather than a blur — stacking offset copies is what
    //    makes the side of a letter read as a side.
    if (sh && (sh.kind === '3d' || sh.kind === 'long')) {
        const depth = Math.max(1, num(sh.depth, 0.08) * sizePx);
        const a = (num(sh.angle, 45) * Math.PI) / 180;
        const steps = Math.max(2, Math.round(depth));
        ctx.fillStyle = sh.color || 'rgba(0,0,0,0.55)';
        for (let i = steps; i >= 1; i--) {
            const k = (i / steps) * depth;
            ctx.save();
            ctx.translate(Math.cos(a) * k, Math.sin(a) * k);
            // A 'long' shadow fades out; a '3d' extrude does not — it is a
            // solid surface, and fading it makes the letter look transparent.
            if (sh.kind === 'long') ctx.globalAlpha = 1 - i / (steps + 1);
            forEachChar(ctx, laid, (ch) => ctx.fillText(ch, 0, 0));
            ctx.restore();
        }
    }

    // 2. Drop shadow / glow, cast by the stroke+fill passes below through canvas
    //    shadow state. Set once, so the shadow comes off the finished letter
    //    rather than off every outline in turn.
    const castShadow = Boolean(sh && (sh.kind === 'drop' || sh.kind === 'glow'));
    if (castShadow) {
        ctx.shadowColor = sh.color || 'rgba(0,0,0,0.45)';
        ctx.shadowBlur = Math.max(0, num(sh.blur, 0.05) * sizePx);
        ctx.shadowOffsetX = num(sh.dx, 0) * sizePx;
        ctx.shadowOffsetY = num(sh.dy, 0) * sizePx;
    }

    // 3. Outlines, widest first and UNDER the fill. Widths are doubled because
    //    half of a canvas stroke falls inside the glyph where the fill covers
    //    it — so an unadjusted width reads as half of what was asked for.
    const strokes = st.strokes
        .filter((s) => s && num(s.width, 0) > 0)
        .sort((a, b) => num(b.width, 0) - num(a.width, 0));
    for (const s of strokes) {
        ctx.strokeStyle = s.color || '#000000';
        ctx.lineWidth = num(s.width, 0.05) * sizePx * 2;
        forEachChar(ctx, laid, (ch) => ctx.strokeText(ch, 0, 0));
        // Only the OUTERMOST pass casts the shadow. The inner outlines sit on
        // top of it, and re-casting from each would build a muddy rim.
        if (castShadow) { ctx.shadowColor = 'transparent'; ctx.shadowBlur = 0; }
    }

    // 4. The fill. With no strokes, this is what casts the shadow.
    ctx.fillStyle = fillStyle(ctx, st.fill, laid.box);
    forEachChar(ctx, laid, (ch) => ctx.fillText(ch, 0, 0));

    ctx.restore();
}

/**
 * The selection box. Measured from the real layout, so a warped headline gets
 * handles around what is on screen rather than around the flat text it would
 * have been.
 */
function wordArtBounds(ctx, plan, layer) {
    if (!layer.text) return null;
    const laid = layoutFor(ctx, layer);
    return { x: layer.x + laid.box.x, y: layer.y + laid.box.y, w: laid.box.w, h: laid.box.h };
}

// ── the stored form ──────────────────────────────────────────────────────────

/**
 * A patch → a compiled word-art layer.
 *
 * Geometry is fractional exactly like a `text` patch: `sizeFrac` is the font
 * size as a fraction of the frame WIDTH, `fx`/`fy` the anchor. That is what
 * lets one recipe render at 9:16 and at 1:1 without a second stored copy, and
 * it is why the image editor gets "this creative at four sizes" for free.
 */
function compileWordArt(patch, cx) {
    const { W, H } = cx;
    const resolved = cx.tokens(String(patch.text == null ? '' : patch.text));
    if (!resolved.text && !patch.keepEmpty) return null;
    const sizeFrac = num(patch.sizeFrac, 0.09);
    const sizePx = Math.max(12, Math.round(W * sizeFrac));
    const preset = PRESETS[patch.preset];
    const style = { ...(preset ? preset.style : {}), ...(patch.style || {}) };
    const env = cx.envelope(patch.anim);

    return {
        name: patch.name || 'Word art',
        text: resolved.text,
        // A tokened headline with nothing behind the token does not draw — the
        // same rule the text layer follows, for the same reason.
        visible: patch.visible !== false && !resolved.missing,
        missingToken: resolved.missing,
        timing: patch.timing || null,
        enter: patch.enter || env.enter,
        exit: patch.exit || env.exit,
        z: patch.z == null ? cx.nextZ() : patch.z,
        x: W * (patch.fx == null ? 0.5 : patch.fx),
        y: H * (patch.fy == null ? 0.5 : patch.fy),
        sizePx,
        style,
        preset: patch.preset || null,
        // Round-trip fractions, so an editor can hand the patch straight back.
        fx: patch.fx == null ? 0.5 : patch.fx,
        fy: patch.fy == null ? 0.5 : patch.fy,
        sizeFrac,
        rot: patch.rot || 0,
        opacity: patch.opacity == null ? 1 : patch.opacity,
        blend: patch.blend || null,
        mask: patch.mask || null,
        anim: patch.anim || 'none',
        keys: cx.keys(patch.keys),
    };
}

export const WORDART = registerLayerType('wordart', {
    paint: paintWordArt,
    compile: compileWordArt,
    bounds: wordArtBounds,
});

/**
 * Resize a word-art layer from a corner drag.
 *
 * Deliberately NOT routed through `resizePatch` in index.js, which handles box
 * layers carrying an `fw`: word art has no stored width — its width is a
 * consequence of the text and the font size. So a corner drag changes
 * `sizeFrac`, exactly as it does for a text layer.
 */
export function wordArtResize(layer, k) {
    const base = layer.sizeFrac || 0.09;
    return { id: layer.id, sizeFrac: +clamp(base * k, 0.02, 0.4).toFixed(4) };
}
