/**
 * The pointer of a screen recording (docs/STUDIO.md R7) and the automatic
 * zoom it drives (X5).
 *
 * While the screen is recorded the main process writes down where the
 * pointer is, a sample a frame, and where it was pressed, as fractions of
 * the recorded screen or window. On the timeline that data is a `pointer`
 * layer in the recording's group, on the lane above the screen: a highlight
 * round the pointer and a ring where it clicked, drawn wherever the screen
 * layer draws that point (through its crop, its fit, an inset's box and its
 * zoom at that instant), so the look can be changed or switched off after,
 * and the screen can be zoomed without the rings falling off their buttons.
 *
 * The automatic zoom reads the clicks: where several land close together in
 * a short time the screen zooms in on them, follows them, and zooms out
 * again. The zooms are kept on the screen layer as a list (`autoZoom`, each
 * { start, end, x, y, level } in the layer's own time), and its zoom and pan
 * keys are made from that list, so a zoom can be made stronger, weaker or
 * taken away after.
 *
 * Pure apart from the painter, which draws with a canvas it is handed.
 */

import { registerLayerType, sourcePointToFrame } from './index.js';
import { clipById, clipsOf } from './project.js';

const round3 = (n) => Math.round(n * 1000) / 1000;

// ── the data ────────────────────────────────────────────────────────────────

/**
 * Pointer data as the recorder hands it over, made small and safe to keep:
 * samples [t, x, y] and clicks [t, x, y], seconds and fractions, rounded to
 * a thousandth; a sample that says nothing new (the pointer resting) is
 * dropped, keeping the last before a move so a move starts where it did,
 * and the very last so a rest runs to the end.
 */
export function compactPointer({ samples = [], clicks = [] } = {}) {
    const ok = (s) => Array.isArray(s) && s.length >= 3 && s.slice(0, 3).every((v) => Number.isFinite(Number(v)));
    const pts = samples.filter(ok).map((s) => [round3(Number(s[0])), round3(Number(s[1])), round3(Number(s[2]))]).sort((a, b) => a[0] - b[0]);
    const out = [];
    for (let i = 0; i < pts.length; i++) {
        const p = pts[i];
        const prev = out[out.length - 1];
        const next = pts[i + 1];
        const same = prev && prev[1] === p[1] && prev[2] === p[2];
        // A resting stretch keeps its first and last sample only.
        // The last sample is kept too, or a pointer resting to the end would vanish at its first rest.
        if (same && next && next[1] === p[1] && next[2] === p[2]) continue;
        out.push(p);
    }
    const hits = clicks.filter(ok).map((c) => [round3(Number(c[0])), round3(Number(c[1])), round3(Number(c[2]))]).sort((a, b) => a[0] - b[0]);
    return { samples: out, clicks: hits };
}

/**
 * Where the pointer is at source time `t`: { x, y }, between the samples
 * either side, or null before the first, after the last, or across a gap of
 * more than half a second (the pointer was off the recorded window).
 */
export function pointerAt(samples, t) {
    if (!Array.isArray(samples) || !samples.length) return null;
    if (t < samples[0][0] - 0.05 || t > samples[samples.length - 1][0] + 0.05) return null;
    let lo = 0;
    let hi = samples.length - 1;
    while (hi - lo > 1) {
        const mid = (lo + hi) >> 1;
        if (samples[mid][0] <= t) lo = mid; else hi = mid;
    }
    const a = samples[lo];
    const b = samples[hi];
    if (t <= a[0] || lo === hi) return { x: a[1], y: a[2] };
    if (t >= b[0]) return { x: b[1], y: b[2] };
    // A resting pointer is written down only where it stopped and where it
    // moved again, so a long gap between two samples at one place is a rest.
    if (b[0] - a[0] > 0.5 && (a[1] !== b[1] || a[2] !== b[2])) return { x: a[1], y: a[2] };
    const k = (t - a[0]) / (b[0] - a[0]);
    return { x: a[1] + (b[1] - a[1]) * k, y: a[2] + (b[2] - a[2]) * k };
}

/** The clicks a ring is showing for at source time `t`, each with how far through its ring it is (0 to 1). */
export function ringsAt(clicks, t, seconds = RING_SECONDS) {
    const out = [];
    for (const c of clicks || []) {
        const age = t - c[0];
        if (age >= 0 && age <= seconds) out.push({ x: c[1], y: c[2], p: age / seconds });
    }
    return out;
}

/** How long a click's ring shows. */
export const RING_SECONDS = 0.6;

/** The pointer layer's look, as it starts: a soft highlight round it and rings where it clicks; the drawn arrow off (the recording has the system's own). */
export const POINTER_LOOK = Object.freeze({ highlight: true, rings: true, arrow: false, size: 1, color: '#ffd54a' });

// ── the layer ───────────────────────────────────────────────────────────────

function compilePointer(patch, cx) {
    return {
        name: patch.name || 'Pointer',
        visible: patch.visible !== false,
        timing: patch.timing || null,
        enter: patch.enter || { kind: 'none', seconds: 0 },
        exit: patch.exit || { kind: 'none', seconds: 0 },
        z: patch.z == null ? cx.nextZ() : patch.z,
        offset: Number(patch.offset) || 0,
        // At its screen recording's speed (setSpeed sets the group together).
        speed: Number(patch.speed) > 0 ? Number(patch.speed) : 1,
        screenUrl: patch.screenUrl || null,
        samples: Array.isArray(patch.samples) ? patch.samples : [],
        clicks: Array.isArray(patch.clicks) ? patch.clicks : [],
        highlight: patch.highlight !== false,
        rings: patch.rings !== false,
        arrow: Boolean(patch.arrow),
        size: Number(patch.size) > 0 ? Number(patch.size) : 1,
        color: patch.color || POINTER_LOOK.color,
        keys: null,
    };
}

/** The screen layer a pointer layer draws over at `time`: the recording's picture of the same file, on the clock there. */
function screenFor(plan, layer, time) {
    return plan.layers.find((l) => l.type === 'video' && l.visible !== false && l.url && l.url === layer.screenUrl
        && (!l.timing || (time >= l.timing.start && time <= l.timing.end))) || null;
}

/** A hex colour with an alpha, for the highlight's soft fill. */
function withAlpha(hex, a) {
    const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || ''));
    if (!m) return `rgba(255,213,74,${a})`;
    const n = parseInt(m[1], 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

/** The arrow a pointer is drawn as, its tip at (x, y), `s` tall. */
function drawArrow(ctx, x, y, s) {
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x, y + s);
    ctx.lineTo(x + s * 0.28, y + s * 0.74);
    ctx.lineTo(x + s * 0.47, y + s * 1.08);
    ctx.lineTo(x + s * 0.6, y + s * 1.02);
    ctx.lineTo(x + s * 0.42, y + s * 0.68);
    ctx.lineTo(x + s * 0.76, y + s * 0.68);
    ctx.closePath();
    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = '#111111';
    ctx.lineWidth = Math.max(1, s * 0.06);
    ctx.fill();
    ctx.stroke();
}

function paintPointer(ctx, plan, layer, time) {
    const screen = screenFor(plan, layer, time);
    if (!screen) return;
    const t = (layer.offset || 0) + (time - (layer.timing ? layer.timing.start : 0)) * (layer.speed || 1);
    const r = plan.W * 0.022 * (layer.size || 1);
    if (layer.rings) {
        for (const ring of ringsAt(layer.clicks, t)) {
            const at = sourcePointToFrame(plan, screen, time, ring.x, ring.y);
            if (!at) continue;
            ctx.save();
            ctx.globalAlpha *= 1 - ring.p;
            ctx.strokeStyle = layer.color;
            ctx.lineWidth = Math.max(2, r * 0.22);
            ctx.beginPath();
            ctx.arc(at.x, at.y, r * (0.55 + ring.p * 1.4), 0, Math.PI * 2);
            ctx.stroke();
            ctx.restore();
        }
    }
    const pos = pointerAt(layer.samples, t);
    if (!pos) return;
    const at = sourcePointToFrame(plan, screen, time, pos.x, pos.y);
    if (!at) return;
    if (layer.highlight) {
        ctx.save();
        ctx.fillStyle = withAlpha(layer.color, 0.32);
        ctx.strokeStyle = withAlpha(layer.color, 0.75);
        ctx.lineWidth = Math.max(1.5, r * 0.08);
        ctx.beginPath();
        ctx.arc(at.x, at.y, r, 0, Math.PI * 2);
        ctx.fill();
        ctx.stroke();
        ctx.restore();
    }
    if (layer.arrow) {
        ctx.save();
        drawArrow(ctx, at.x, at.y, r * 1.3);
        ctx.restore();
    }
}

export const POINTER = registerLayerType('pointer', {
    paint: paintPointer,
    compile: compilePointer,
    // The whole frame: a pointer is never dragged on the stage.
    bounds: (ctx, plan) => ({ x: 0, y: 0, w: plan.W, h: plan.H }),
});

// ── the automatic zoom (X5) ─────────────────────────────────────────────────

/** How strongly the automatic zoom zooms, and how long it takes to zoom in and out. */
export const ZOOM_LEVEL = 1.8;
export const ZOOM_RAMP = 0.6;

/**
 * Where the screen should zoom, from the clicks inside a layer's window: each
 * run of clicks close together (under 3 seconds apart, inside a quarter of
 * the screen of where the run began) is a zoom from a little before its
 * first click to a while after its last, on the middle of its clicks. Zooms
 * that would touch are one zoom. Times are the layer's own (0 at its start).
 * Answers [{ start, end, x, y, level, points: [[t, x, y]] }].
 */
export function findZooms(clicks, { offset = 0, length = Infinity, level = ZOOM_LEVEL, before = 0.8, after = 1.6, gap = 3, reach = 0.25 } = {}) {
    const inside = (clicks || [])
        .map((c) => [c[0] - offset, c[1], c[2]])
        .filter((c) => c[0] >= 0 && c[0] <= length)
        .sort((a, b) => a[0] - b[0]);
    const runs = [];
    for (const c of inside) {
        const run = runs[runs.length - 1];
        if (run && c[0] - run.points[run.points.length - 1][0] <= gap && Math.hypot(c[1] - run.points[0][1], c[2] - run.points[0][2]) <= reach) run.points.push(c);
        else runs.push({ points: [c] });
    }
    const zooms = [];
    for (const run of runs) {
        const first = run.points[0][0];
        const last = run.points[run.points.length - 1][0];
        const x = run.points.reduce((s, p) => s + p[1], 0) / run.points.length;
        const y = run.points.reduce((s, p) => s + p[2], 0) / run.points.length;
        const z = { start: round3(Math.max(0, first - before)), end: round3(Math.min(length, last + after)), x: round3(x), y: round3(y), level, points: run.points.map((p) => p.map(round3)) };
        const prev = zooms[zooms.length - 1];
        if (prev && z.start - prev.end < ZOOM_RAMP * 2) {
            const all = [...prev.points, ...z.points];
            prev.end = z.end;
            prev.points = all;
            prev.x = round3(all.reduce((s, p) => s + p[1], 0) / all.length);
            prev.y = round3(all.reduce((s, p) => s + p[2], 0) / all.length);
        } else zooms.push(z);
    }
    return zooms.filter((z) => z.end - z.start > 0.3);
}

/** A point of the source as the screen layer's pan: fractions of its crop, so the zoom stays inside it. */
function panOf(x, y, crop) {
    const c = crop || { x: 0, y: 0, w: 1, h: 1 };
    const clamp = (v) => Math.max(0, Math.min(1, v));
    return { panX: round3(clamp((x - c.x) / (c.w || 1))), panY: round3(clamp((y - c.y) / (c.h || 1))) };
}

/**
 * The zoom and pan keys a list of zooms makes on a layer of `length`
 * seconds: zoom 1 until a zoom ramps in, held at its level while the pan
 * follows its clicks, and ramped out again. Keys are the layer's own time.
 */
export function zoomKeys(zooms, { crop = null, ramp = ZOOM_RAMP } = {}) {
    const zoom = [];
    const panX = [];
    const panY = [];
    const sorted = [...(zooms || [])].sort((a, b) => a.start - b.start);
    for (const z of sorted) {
        const level = Math.max(1, Number(z.level) || ZOOM_LEVEL);
        const r = Math.min(ramp, (z.end - z.start) / 3);
        const inAt = round3(z.start + r);
        const outAt = round3(z.end - r);
        zoom.push({ t: round3(z.start), v: 1 }, { t: inAt, v: level, ease: 'in-out' }, { t: outAt, v: level }, { t: round3(z.end), v: 1, ease: 'in-out' });
        // The pan: the middle of the run while it zooms in, then each click in turn while it holds.
        const centre = panOf(z.x, z.y, crop);
        panX.push({ t: round3(z.start), v: centre.panX }, { t: inAt, v: centre.panX });
        panY.push({ t: round3(z.start), v: centre.panY }, { t: inAt, v: centre.panY });
        const pts = (z.points || []).filter((p) => p[0] > inAt + 0.2 && p[0] < outAt - 0.2);
        for (const p of pts) {
            const at = panOf(p[1], p[2], crop);
            panX.push({ t: round3(p[0]), v: at.panX, ease: 'in-out' });
            panY.push({ t: round3(p[0]), v: at.panY, ease: 'in-out' });
        }
        const last = pts.length ? panOf(pts[pts.length - 1][1], pts[pts.length - 1][2], crop) : centre;
        panX.push({ t: round3(z.end), v: last.panX });
        panY.push({ t: round3(z.end), v: last.panY });
    }
    return zoom.length ? { zoom, panX, panY } : null;
}

/** The pointer layer that belongs to a screen layer: the same recording's, over the same file. */
export function pointerOf(doc, screenId) {
    const screen = clipById(doc, screenId);
    if (!screen) return null;
    return clipsOf(doc).find((p) => p.type === 'pointer' && p.screenUrl === screen.url
        && (!screen.group || p.group === screen.group)
        && p.timing.start < screen.timing.end && p.timing.end > screen.timing.start) || null;
}

/** A screen layer's keys with the zoom's own taken out (any other motion kept), then the list's put in. */
function withZoomKeys(layer, zooms) {
    const keys = { ...(layer.keys || {}) };
    delete keys.zoom;
    delete keys.panX;
    delete keys.panY;
    const made = zoomKeys(zooms, { crop: layer.crop });
    const next = { ...keys, ...(made || {}) };
    return Object.keys(next).length ? next : undefined;
}

/**
 * Zoom a screen layer automatically from its recording's clicks: the zooms
 * found are kept on it as `autoZoom` and made into its keys. `level` sets
 * how strongly. Answers the new document, the same one when there are no
 * clicks to zoom on.
 */
export function applyAutoZoom(doc, screenId, { level = ZOOM_LEVEL } = {}) {
    const screen = clipById(doc, screenId);
    const ptr = pointerOf(doc, screenId);
    if (!screen || !ptr) return doc;
    // The pointer's times are the recording's, as the screen file's are: a click at t is the screen layer's t less its offset.
    const zooms = findZooms(ptr.clicks, { offset: screen.offset || 0, length: screen.timing.end - screen.timing.start, level });
    return setZooms(doc, screenId, zooms);
}

/** Put a list of zooms on a screen layer (an empty list takes the automatic zoom off). */
export function setZooms(doc, screenId, zooms) {
    const list = (zooms || []).filter((z) => z && z.end > z.start);
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (p.id !== screenId) return p;
            const next = { ...p, autoZoom: list.length ? list : undefined, keys: withZoomKeys(p, list) };
            if (!next.autoZoom) delete next.autoZoom;
            if (!next.keys) delete next.keys;
            return next;
        }),
    };
}

/** One zoom changed (its level, or anything else) or, with `null`, taken away; the keys follow. */
export function updateZoom(doc, screenId, index, change) {
    const screen = clipById(doc, screenId);
    if (!screen?.autoZoom) return doc;
    const list = screen.autoZoom.map((z, i) => (i === index ? (change === null ? null : { ...z, ...change }) : z)).filter(Boolean);
    return setZooms(doc, screenId, list);
}
