/**
 * The two renderer facts the editor's paint scheduling is built on.
 *
 * Neither is a new behaviour — both have been true since the lift — but the
 * editor now DEPENDS on them, and a change to either would come back as
 * flickering rather than as a failing assertion. So they are pinned here, with
 * the consequence written down next to each.
 *
 *   node --test tests/painting.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { buildPlan, paintFrame } from '../packages/studio/src/index.js';

/** A 2D context that records what was done to it. */
function stubCtx() {
    const calls = [];
    const state = {};
    const record = (name) => (...args) => { calls.push({ name, args }); };
    const ctx = {
        calls,
        save: record('save'), restore: record('restore'),
        translate: record('translate'), rotate: record('rotate'), scale: record('scale'),
        beginPath: record('beginPath'), moveTo: record('moveTo'), lineTo: record('lineTo'),
        quadraticCurveTo: record('quadraticCurveTo'),
        bezierCurveTo: record('bezierCurveTo'), closePath: record('closePath'),
        ellipse: record('ellipse'), rect: record('rect'), roundRect: record('roundRect'),
        clip: record('clip'), fill: record('fill'), stroke: record('stroke'),
        fillRect: record('fillRect'), clearRect: record('clearRect'),
        setLineDash: record('setLineDash'),
        fillText: record('fillText'), strokeText: record('strokeText'),
        drawImage: record('drawImage'), putImageData: record('putImageData'),
        measureText: (s) => ({ width: String(s).length * 10 }),
        createLinearGradient: () => ({ addColorStop() {} }),
        createRadialGradient: () => ({ addColorStop() {} }),
        createPattern: () => null,
        getImageData: (x, y, w, h) => {
            calls.push({ name: 'getImageData', args: [x, y, w, h] });
            return { data: new Uint8ClampedArray(w * h * 4).fill(128), width: w, height: h };
        },
    };
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

/**
 * A canvas that records every write to `width` / `height`.
 *
 * The browser's own canvas resets its bitmap to transparent when either is
 * assigned — even to the value it already holds — so every write recorded here
 * is a frame gone blank on a real one.
 */
function recordingCanvas() {
    const writes = [];
    let w = 300;
    let h = 150;
    const ctx = stubCtx();
    return {
        writes,
        getContext: () => ctx,
        get width() { return w; },
        set width(v) { writes.push(['width', v]); w = v; },
        get height() { return h; },
        set height(v) { writes.push(['height', v]); h = v; },
    };
}

const shared = (canvas) => ({
    canvas,
    images: [],
    title: 'Title',
    body: 'Some words for the caption to lay out and type.',
    options: { aspect: 'vertical', theme: 'dark' },
    layerPatches: [],
});

// ── why the editor compiles against a scratch canvas ────────────────────────

test('buildPlan writes the canvas dimensions it is handed — every time', () => {
    // THIS IS THE FLICKER, at its source. On a real canvas each of these writes
    // blanks the bitmap, and the editor recompiles on every drag tick. That is
    // why useCreative compiles against a scratch canvas nobody is looking at
    // and paints the visible one separately.
    //
    // If this ever starts failing because buildPlan grew a guard, the scratch
    // canvas can be simplified — but read the note on scratchRef first, because
    // the passive-effect half of the bug is a separate matter.
    const canvas = recordingCanvas();
    buildPlan(shared(canvas));
    assert.ok(canvas.writes.length > 0, 'buildPlan sized the canvas');
    assert.ok(canvas.writes.some(([k]) => k === 'width'), 'including its width, which is what clears it');
});

test('compiling twice at an unchanged size still writes the dimensions', () => {
    // The property that makes it a per-tick clear rather than a one-off: no
    // guard, so an unchanged size costs exactly as much as a changed one.
    const canvas = recordingCanvas();
    buildPlan(shared(canvas));
    const first = canvas.writes.length;
    buildPlan(shared(canvas));
    assert.ok(canvas.writes.length > first, 'a second identical compile wrote again');
});

test('a plan is not bound to the canvas it was compiled on', () => {
    // What makes the scratch canvas sound. paintFrame takes its target
    // explicitly, so a plan compiled on one canvas paints correctly onto
    // another of the same size — which the deck presenter already relies on.
    const compiled = recordingCanvas();
    const plan = buildPlan(shared(compiled));
    const target = stubCtx();
    paintFrame(target, plan, 0);
    assert.ok(target.calls.some((c) => c.name === 'fillRect'), 'it painted onto the other context');
});

// ── why the visible canvas never needs clearing ─────────────────────────────

test('paintFrame fills the whole frame before it draws anything', () => {
    // The other half. Because the first paint op covers the entire frame
    // opaquely, the display canvas never has to be cleared between frames — so
    // nothing in the editor needs to blank it, and nothing should.
    const canvas = recordingCanvas();
    const plan = buildPlan(shared(canvas));
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);

    const firstDraw = ctx.calls.findIndex((c) => c.name === 'fillRect' || c.name === 'drawImage' || c.name === 'fillText');
    assert.ok(firstDraw >= 0, 'it drew something');
    const first = ctx.calls[firstDraw];
    assert.equal(first.name, 'fillRect', 'the first thing painted is the background');
    assert.deepEqual(first.args, [0, 0, plan.W, plan.H], 'and it covers the entire frame');
});

test('paintFrame never clears the canvas it paints onto', () => {
    // clearRect would leave a transparent frame if anything after it failed,
    // and it is unnecessary given the opaque fill above.
    const canvas = recordingCanvas();
    const plan = buildPlan(shared(canvas));
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    assert.ok(!ctx.calls.some((c) => c.name === 'clearRect'), 'nothing was cleared');
});

test('painting does not resize the canvas it paints onto', () => {
    // The invariant the editor's display canvas depends on: React sizes it from
    // `frame`, and painting must never touch that — a resize mid-gesture is a
    // blank frame.
    const compiled = recordingCanvas();
    const plan = buildPlan(shared(compiled));

    const target = recordingCanvas();
    const before = target.writes.length;
    paintFrame(target.getContext('2d'), plan, 0);
    assert.equal(target.writes.length, before, 'paintFrame wrote no dimensions');
});
