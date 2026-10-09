/**
 * The Studio toolset — word art, shapes, adjustment layers.
 *
 * These run in Node with no browser, which is the point: the renderer's own
 * contract is that a plan is data and painting is pure, so everything except
 * the actual rasterisation can be checked without a canvas. What cannot be
 * checked here — do the pixels look right — is what the browser harness
 * (`npm run gate`) is for.
 *
 * The context stub RECORDS calls rather than pretending to draw. That is enough
 * to assert the things that actually break in this code: that outlines are
 * painted under the fill and widest-first, that a warp rotates characters, that
 * an adjustment goes back through drawImage rather than putImageData, and that
 * the extraction gate's invariant still holds — a plan carrying none of these
 * layers touches none of this code.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { applyLayerPatches, paintFrame, layerBounds, hitTestLayers, registeredLayerTypes } from '../packages/studio/src/index.js';
import { applyAdjustment, ADD_MENU, WORDART_PRESETS, GEOMETRIES } from '../packages/studio/src/toolset.js';

// ── a recording 2D context ───────────────────────────────────────────────────

function stubCtx() {
    const calls = [];
    const state = {};
    const record = (name) => (...args) => { calls.push({ name, args, lineWidth: state.lineWidth, strokeStyle: state.strokeStyle }); };
    const ctx = {
        calls,
        save: record('save'),
        restore: record('restore'),
        translate: record('translate'),
        rotate: record('rotate'),
        scale: record('scale'),
        beginPath: record('beginPath'),
        moveTo: record('moveTo'),
        lineTo: record('lineTo'),
        quadraticCurveTo: record('quadraticCurveTo'),
        bezierCurveTo: record('bezierCurveTo'),
        closePath: record('closePath'),
        ellipse: record('ellipse'),
        rect: record('rect'),
        roundRect: record('roundRect'),
        clip: record('clip'),
        fill: record('fill'),
        stroke: record('stroke'),
        fillRect: record('fillRect'),
        setLineDash: record('setLineDash'),
        fillText: record('fillText'),
        strokeText: record('strokeText'),
        drawImage: record('drawImage'),
        putImageData: record('putImageData'),
        // 10px per character is wrong but stable, which is all a layout test needs.
        measureText: (s) => ({ width: String(s).length * 10 }),
        createLinearGradient: () => ({ addColorStop() {} }),
        createRadialGradient: () => ({ addColorStop() {} }),
        // Recorded AND returns pixels: an adjustment layer needs both, and a
        // stub that only did one of them cost a debugging session.
        getImageData: (x, y, w, h) => {
            calls.push({ name: 'getImageData', args: [x, y, w, h] });
            return { data: new Uint8ClampedArray(w * h * 4).fill(128), width: w, height: h };
        },
    };
    // Track the settable properties the painters use, so assertions can read them.
    for (const prop of ['font', 'fillStyle', 'strokeStyle', 'lineWidth', 'globalAlpha', 'filter',
        'globalCompositeOperation', 'shadowColor', 'shadowBlur', 'shadowOffsetX', 'shadowOffsetY',
        'textAlign', 'textBaseline', 'lineJoin', 'miterLimit', 'direction']) {
        Object.defineProperty(ctx, prop, {
            get: () => state[prop],
            set: (v) => { state[prop] = v; calls.push({ name: `set:${prop}`, args: [v] }); },
        });
    }
    return ctx;
}

/** The smallest object applyLayerPatches and paintFrame will accept as a plan. */
function stubPlan(overrides = {}) {
    return {
        W: 1080, H: 1920, duration: 5, fps: 30,
        theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
        layers: [], images: [], videos: {}, assets: {}, context: {},
        ...overrides,
    };
}

const compile = (patches, planOverrides) => {
    const plan = stubPlan(planOverrides);
    applyLayerPatches(plan, patches);
    return plan;
};

// ── registration ─────────────────────────────────────────────────────────────

test('the three toolset types register themselves on import', () => {
    const types = registeredLayerTypes();
    for (const t of ['wordart', 'shape', 'adjust']) assert.ok(types.includes(t), `${t} not registered`);
});

test('every Add-menu entry names a type that is actually registered', () => {
    const types = registeredLayerTypes();
    for (const entry of ADD_MENU) {
        assert.ok(types.includes(entry.type), `Add menu offers unregistered type ${entry.type}`);
        assert.equal(entry.patch.type, entry.type, 'the patch must build the type the entry claims');
    }
});

// ── the extraction gate's invariant ──────────────────────────────────────────

test('a plan with no toolset layers touches no toolset code', () => {
    // The whole basis for extending the renderer in place: the additions are
    // guarded on properties a baseline plan does not carry. If this ever fails,
    // the A/B gate is about to go red for a reason nobody will enjoy finding.
    const plan = stubPlan({
        layers: [{ id: 'footer', type: 'text', text: 'x', font: '20px sans-serif', color: '#fff', x: 10, y: 10, z: 1 }],
    });
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const touched = ctx.calls.filter((c) => c.name === 'set:globalCompositeOperation' || c.name === 'clip');
    assert.equal(touched.length, 0, 'blend/mask state was written for a plan that asked for neither');
});

// ── word art ─────────────────────────────────────────────────────────────────

test('word art compiles with fractional geometry it can round-trip', () => {
    const plan = compile([{ id: 'wa', type: 'wordart', text: 'SALE', fx: 0.25, fy: 0.6, sizeFrac: 0.12 }]);
    const l = plan.layers[0];
    assert.equal(l.type, 'wordart');
    assert.equal(l.text, 'SALE');
    assert.equal(l.x, 1080 * 0.25);
    assert.equal(l.y, 1920 * 0.6);
    assert.equal(l.sizePx, Math.round(1080 * 0.12));
    // The fractions ride along, or an editor cannot hand the patch back.
    assert.equal(l.fx, 0.25);
    assert.equal(l.sizeFrac, 0.12);
});

test('a preset supplies the style and an explicit style overrides it', () => {
    const plan = compile([
        { id: 'a', type: 'wordart', text: 'A', preset: 'gold' },
        { id: 'b', type: 'wordart', text: 'B', preset: 'gold', style: { warp: 'arc' } },
    ]);
    assert.equal(plan.layers[0].style.fill.kind, 'linear', 'the gold preset is a gradient');
    assert.equal(plan.layers[1].style.warp, 'arc', 'the explicit style wins');
    assert.equal(plan.layers[1].style.fill.kind, 'linear', 'and does not wipe the rest of the preset');
});

test('an unresolved token hides the headline rather than printing a hole', () => {
    const plan = compile([{ id: 'p', type: 'wordart', text: 'Only {price} today' }], { context: {} });
    assert.equal(plan.layers[0].visible, false);
    assert.equal(plan.layers[0].missingToken, true);
});

test('a resolved token is substituted and the layer draws', () => {
    const plan = compile([{ id: 'p', type: 'wordart', text: 'Only {price} today' }], { context: { price: 'Rs 500' } });
    assert.equal(plan.layers[0].text, 'Only Rs 500 today');
    assert.equal(plan.layers[0].visible, true);
});

test('outlines paint under the fill, widest first', () => {
    const plan = compile([{
        id: 'wa', type: 'wordart', text: 'AB',
        style: { strokes: [{ width: 0.05, color: '#111' }, { width: 0.15, color: '#222' }] },
    }]);
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);

    const strokeWidths = ctx.calls.filter((c) => c.name === 'strokeText').map((c) => c.lineWidth);
    assert.ok(strokeWidths.length >= 4, 'both outlines painted, once per character');
    assert.ok(strokeWidths[0] > strokeWidths[strokeWidths.length - 1], 'widest outline is painted first');

    const firstFill = ctx.calls.findIndex((c) => c.name === 'fillText');
    const lastStroke = ctx.calls.map((c) => c.name).lastIndexOf('strokeText');
    assert.ok(lastStroke < firstFill, 'every outline is under the fill');
});

test('an arc warp rotates characters; no warp rotates nothing', () => {
    const straight = compile([{ id: 'a', type: 'wordart', text: 'ABCDE' }]);
    const arced = compile([{ id: 'a', type: 'wordart', text: 'ABCDE', style: { warp: 'arc', curve: 0.8 } }]);

    const rotations = (plan) => {
        const ctx = stubCtx();
        paintFrame(ctx, plan, 0);
        return ctx.calls.filter((c) => c.name === 'rotate').map((c) => c.args[0]).filter((r) => r !== 0);
    };
    assert.equal(rotations(straight).length, 0, 'a flat headline rotates no characters');
    assert.ok(rotations(arced).length >= 4, 'an arc rotates each character to the tangent');
});

test('word-art bounds follow the warp rather than the flat text', () => {
    const ctx = stubCtx();
    const flat = compile([{ id: 'a', type: 'wordart', text: 'ABCDEFGH', sizeFrac: 0.1 }]);
    const arced = compile([{ id: 'a', type: 'wordart', text: 'ABCDEFGH', sizeFrac: 0.1, style: { warp: 'arc', curve: 0.9 } }]);
    const bFlat = layerBounds(ctx, flat, flat.layers[0]);
    const bArc = layerBounds(ctx, arced, arced.layers[0]);
    assert.ok(bArc.h > bFlat.h, 'a deep arc is taller than the flat line it came from');
});

test('word art is selectable — the registry feeds the hit test', () => {
    const plan = compile([{ id: 'wa', type: 'wordart', text: 'HELLO', fx: 0.5, fy: 0.5, sizeFrac: 0.1 }]);
    const ctx = stubCtx();
    const hit = hitTestLayers(ctx, plan, 1080 * 0.5, 1920 * 0.5);
    assert.ok(hit, 'a click in the middle of the headline selects it');
    // hitTestLayers answers { layer, bounds } — the caller wants the box too.
    assert.equal(hit.layer.id, 'wa');
});

// ── shapes ───────────────────────────────────────────────────────────────────

test('every advertised geometry builds a path without throwing', () => {
    for (const geometry of GEOMETRIES) {
        const plan = compile([{ id: 's', type: 'shape', geometry, fw: 0.4, fh: 0.25 }]);
        const ctx = stubCtx();
        paintFrame(ctx, plan, 0);
        assert.ok(ctx.calls.some((c) => c.name === 'fill'), `${geometry} painted nothing`);
    }
});

test('an unknown geometry degrades to a box instead of vanishing', () => {
    const plan = compile([{ id: 's', type: 'shape', geometry: 'dodecahedron', fw: 0.4, fh: 0.25 }]);
    assert.equal(plan.layers[0].geometry, 'rect');
});

test('fill: null means outline-only and survives compilation', () => {
    const plan = compile([{ id: 's', type: 'shape', fill: null, stroke: { width: 0.01, color: '#fff' } }]);
    assert.equal(plan.layers[0].fill, null, 'a default must not swallow a deliberate null');
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    assert.ok(!ctx.calls.some((c) => c.name === 'fill'), 'nothing was filled');
    assert.ok(ctx.calls.some((c) => c.name === 'stroke'), 'the outline was drawn');
});

test('shape bounds and the painted box agree', () => {
    const plan = compile([{ id: 's', type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 0.4, fh: 0.2 }]);
    const b = layerBounds(stubCtx(), plan, plan.layers[0]);
    assert.equal(b.w, 1080 * 0.4);
    assert.equal(b.h, 1920 * 0.2);
    assert.equal(b.x, 1080 * 0.5 - b.w / 2, 'the box is centred on the anchor');
});

// ── blend and mask, from the frame wrapper ───────────────────────────────────

test('blend and mask are applied by the wrapper for any layer type', () => {
    const plan = compile([{
        id: 's', type: 'shape', blend: 'multiply',
        mask: { shape: 'ellipse', fx: 0.5, fy: 0.5, fw: 0.5 },
    }]);
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    assert.ok(ctx.calls.some((c) => c.name === 'set:globalCompositeOperation' && c.args[0] === 'multiply'));
    assert.ok(ctx.calls.some((c) => c.name === 'clip'), 'the mask clipped');
    assert.ok(ctx.calls.some((c) => c.name === 'ellipse'), 'and it clipped to the shape asked for');
});

test('an inverted mask clips even-odd, which is what punches a hole', () => {
    const plan = compile([{ id: 's', type: 'shape', mask: { shape: 'rect', invert: true, fw: 0.3 } }]);
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const clip = ctx.calls.find((c) => c.name === 'clip');
    assert.equal(clip.args[0], 'evenodd');
});

// ── adjustments ──────────────────────────────────────────────────────────────

/** n pixels of one flat colour, as the raw RGBA an adjustment receives. */
function pixels(r, g, b, count = 4) {
    const d = new Uint8ClampedArray(count * 4);
    for (let i = 0; i < count; i++) { d[i * 4] = r; d[i * 4 + 1] = g; d[i * 4 + 2] = b; d[i * 4 + 3] = 255; }
    return d;
}

test('an empty adjustment is the identity', () => {
    const d = pixels(10, 120, 250);
    const out = applyAdjustment(d, 2, 2, {});
    assert.deepEqual([...out.slice(0, 3)], [10, 120, 250]);
});

test('invert inverts and leaves alpha alone', () => {
    const out = applyAdjustment(pixels(0, 128, 255), 2, 2, { invert: true });
    assert.deepEqual([...out.slice(0, 4)], [255, 127, 0, 255]);
});

test('saturation -1 collapses to luma', () => {
    const out = applyAdjustment(pixels(255, 0, 0), 2, 2, { saturation: -1 });
    const [r, g, b] = out;
    assert.equal(r, g);
    assert.equal(g, b);
    assert.equal(r, Math.round(0.2126 * 255), 'Rec.709 luma, not an average');
});

test('vibrance leaves an already-saturated pixel nearly alone but lifts a dull one', () => {
    const vivid = applyAdjustment(pixels(255, 0, 0), 2, 2, { vibrance: 1 });
    const dull = applyAdjustment(pixels(140, 120, 120), 2, 2, { vibrance: 1 });
    const spread = (d) => Math.max(d[0], d[1], d[2]) - Math.min(d[0], d[1], d[2]);
    assert.equal(spread(vivid), 255, 'a fully saturated pixel cannot get more saturated');
    assert.ok(spread(dull) > 20, 'a dull pixel is pushed well past its original 20-point spread');
});

test('levels move the black and white points', () => {
    // Everything at or below the input black point becomes the output black.
    const out = applyAdjustment(pixels(40, 40, 40), 2, 2, { levels: { inBlack: 60, inWhite: 200 } });
    assert.equal(out[0], 0);
    const hi = applyAdjustment(pixels(220, 220, 220), 2, 2, { levels: { inBlack: 60, inWhite: 200 } });
    assert.equal(hi[0], 255);
});

test('a tone curve is honoured, and its default is a straight line', () => {
    const straight = applyAdjustment(pixels(100, 100, 100), 2, 2, { curves: { rgb: [[0, 0], [255, 255]] } });
    assert.equal(straight[0], 100);
    const lifted = applyAdjustment(pixels(100, 100, 100), 2, 2, { curves: { rgb: [[0, 0], [128, 200], [255, 255]] } });
    assert.ok(lifted[0] > 140, 'the midtone was lifted toward the control point');
});

test('temperature is a per-channel gain, warm meaning more red and less blue', () => {
    const warm = applyAdjustment(pixels(128, 128, 128), 2, 2, { temperature: 0.5 });
    assert.ok(warm[0] > 128, 'red up');
    assert.ok(warm[2] < 128, 'blue down');
});

test('posterize quantises to the number of steps asked for', () => {
    const out = applyAdjustment(pixels(130, 130, 130), 2, 2, { posterize: 2 });
    assert.ok(out[0] === 0 || out[0] === 255, 'two steps means black or white');
});

test('an adjustment layer returns its pixels through drawImage, not putImageData', () => {
    // The distinction is the whole reason for the offscreen surface: only
    // drawImage honours the opacity, mask and blend the wrapper set up.
    globalThis.OffscreenCanvas = class {
        constructor(w, h) { this.width = w; this.height = h; }
        getContext() {
            return {
                putImageData() {}, save() {}, restore() {}, fillRect() {},
                createRadialGradient: () => ({ addColorStop() {} }),
                set globalCompositeOperation(_v) {}, set fillStyle(_v) {},
            };
        }
    };
    try {
        const plan = compile([{ id: 'adj', type: 'adjust', preset: 'punch' }]);
        const ctx = stubCtx();
        paintFrame(ctx, plan, 0);
        assert.ok(ctx.calls.some((c) => c.name === 'getImageData'), 'it read what was beneath it');
        assert.ok(ctx.calls.some((c) => c.name === 'drawImage'), 'and wrote back through drawImage');
        assert.ok(!ctx.calls.some((c) => c.name === 'putImageData'), 'never straight onto the frame');
    } finally {
        delete globalThis.OffscreenCanvas;
    }
});

test('a full-frame adjustment defaults to the whole frame', () => {
    const plan = compile([{ id: 'adj', type: 'adjust' }]);
    const b = layerBounds(stubCtx(), plan, plan.layers[0]);
    assert.deepEqual(b, { x: 0, y: 0, w: 1080, h: 1920 });
});

test('a preset merges under an explicit adjust block', () => {
    const plan = compile([{ id: 'adj', type: 'adjust', preset: 'noir', adjust: { vignette: 0.1 } }]);
    assert.equal(plan.layers[0].adjust.saturation, -1, 'kept from the preset');
    assert.equal(plan.layers[0].adjust.vignette, 0.1, 'overridden explicitly');
});

test('word-art presets are all well formed', () => {
    for (const [key, preset] of Object.entries(WORDART_PRESETS)) {
        assert.ok(preset.label, `${key} has no label`);
        assert.equal(typeof preset.style, 'object', `${key} has no style`);
    }
});
