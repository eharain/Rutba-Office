/**
 * Text blocks — wrapping, lists, auto-fit.
 *
 * All of it is layout, which is exactly the kind of thing that looks right in a
 * screenshot and is wrong in the case nobody tried: the long word, the blank
 * line between items, the box that is one line too short. So the assertions are
 * about the LAID-OUT lines rather than about pixels.
 *
 * The stub context measures 10px per character, which is wrong but stable —
 * all a layout test needs, and the same convention toolset.test.mjs uses.
 *
 *   node --test tests/textblock.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { applyLayerPatches, paintFrame, layerBounds, registeredLayerTypes } from '../packages/studio/src/index.js';
import { LIST_KINDS, LIST_LABELS } from '../packages/studio/src/textblock.js';

const CHAR = 10;

function stubCtx() {
    const calls = [];
    const state = {};
    const record = (name) => (...args) => { calls.push({ name, args, font: state.font, fillStyle: state.fillStyle }); };
    const ctx = {
        calls,
        save: record('save'), restore: record('restore'),
        translate: record('translate'), rotate: record('rotate'), scale: record('scale'),
        beginPath: record('beginPath'), moveTo: record('moveTo'), lineTo: record('lineTo'),
        quadraticCurveTo: record('quadraticCurveTo'), bezierCurveTo: record('bezierCurveTo'),
        closePath: record('closePath'), ellipse: record('ellipse'), rect: record('rect'),
        roundRect: record('roundRect'), clip: record('clip'), fill: record('fill'), stroke: record('stroke'),
        fillRect: record('fillRect'), setLineDash: record('setLineDash'),
        fillText: record('fillText'), strokeText: record('strokeText'),
        drawImage: record('drawImage'), putImageData: record('putImageData'),
        measureText: (s) => ({ width: String(s).length * CHAR }),
        createLinearGradient: () => ({ addColorStop() {} }),
        createRadialGradient: () => ({ addColorStop() {} }),
        getImageData: (x, y, w, h) => ({ data: new Uint8ClampedArray(w * h * 4), width: w, height: h }),
    };
    for (const p of ['font', 'fillStyle', 'strokeStyle', 'lineWidth', 'globalAlpha', 'filter',
        'globalCompositeOperation', 'shadowColor', 'shadowBlur', 'shadowOffsetX', 'shadowOffsetY',
        'textAlign', 'textBaseline', 'lineJoin', 'lineCap', 'miterLimit', 'direction']) {
        Object.defineProperty(ctx, p, { get: () => state[p], set: (v) => { state[p] = v; } });
    }
    return ctx;
}

const stubPlan = (over = {}) => ({
    W: 1000, H: 1000, duration: 5, fps: 30,
    theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
    layers: [], images: [], videos: {}, assets: {}, context: {}, ...over,
});

const compile = (patch, over) => {
    const plan = stubPlan(over);
    applyLayerPatches(plan, [{ id: 'tb', type: 'textblock', ...patch }]);
    return plan;
};

/** The text drawn, line by line, in order. */
function drawnLines(patch, over) {
    const plan = compile(patch, over);
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    return ctx.calls.filter((c) => c.name === 'fillText').map((c) => c.args[0]);
}

// ── registration ────────────────────────────────────────────────────────────

test('the text block registers itself on import', () => {
    assert.ok(registeredLayerTypes().includes('textblock'));
});

test('every list kind has a label', () => {
    for (const k of LIST_KINDS) assert.ok(LIST_LABELS[k], `${k} has no label`);
});

// ── wrapping ────────────────────────────────────────────────────────────────

test('text wraps to the box rather than running out of it', () => {
    // fw 0.4 of a 1000px frame = 400px = 40 characters at the stub's 10px each.
    const lines = drawnLines({ text: 'aaaa '.repeat(30).trim(), fw: 0.4, sizeFrac: 0.02, autoFit: false });
    assert.ok(lines.length > 1, 'it wrapped');
    for (const l of lines) assert.ok(l.length * CHAR <= 400, `"${l}" is wider than the box`);
});

test('a word longer than the line is broken rather than allowed to overflow', () => {
    // The URL case. Without this the line simply runs off the frame.
    const long = 'x'.repeat(120);
    const lines = drawnLines({ text: long, fw: 0.3, sizeFrac: 0.02, autoFit: false });
    assert.ok(lines.length > 1, 'the long word was split');
    for (const l of lines) assert.ok(l.length * CHAR <= 300, 'a fragment still overflows');
});

test('an explicit newline starts a new line', () => {
    const lines = drawnLines({ text: 'one\ntwo', fw: 0.9, sizeFrac: 0.02, autoFit: false });
    assert.deepEqual(lines, ['one', 'two']);
});

test('a blank line between paragraphs is kept', () => {
    // It is spacing somebody typed on purpose; collapsing it silently re-flows
    // their layout.
    const plan = compile({ text: 'one\n\ntwo', fw: 0.9, sizeFrac: 0.02, autoFit: false });
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const ys = ctx.calls.filter((c) => c.name === 'fillText').map((c) => c.args[2]);
    assert.equal(ys.length, 2, 'the blank line draws nothing');
    // ...but it still advances, so the two visible lines are two leads apart.
    assert.ok(ys[1] - ys[0] > 1, 'the gap survived');
});

// ── lists ───────────────────────────────────────────────────────────────────

test('a bulleted list marks each item once', () => {
    const drawn = drawnLines({ text: 'alpha\nbeta\ngamma', list: 'bullet', fw: 0.9, sizeFrac: 0.02, autoFit: false });
    assert.equal(drawn.filter((t) => t === '•').length, 3);
});

test('a numbered list counts items, not lines', () => {
    // THE REGRESSION THIS GUARDS: a wrapped item must not take the next number.
    const text = `${'w '.repeat(40).trim()}\nshort`;
    const drawn = drawnLines({ text, list: 'number', fw: 0.3, sizeFrac: 0.02, autoFit: false });
    const markers = drawn.filter((t) => /^\d+\.$/.test(t));
    assert.deepEqual(markers, ['1.', '2.'], 'the wrapped first item is still item one');
});

test('a wrapped list item hangs under its own text, with no second marker', () => {
    const text = 'w '.repeat(40).trim();
    const drawn = drawnLines({ text, list: 'bullet', fw: 0.3, sizeFrac: 0.02, autoFit: false });
    assert.equal(drawn.filter((t) => t === '•').length, 1, 'one item, one bullet');
    assert.ok(drawn.length > 2, 'and it did wrap');
});

test('a blank line in a list is not a bulleted nothing', () => {
    const drawn = drawnLines({ text: 'alpha\n\nbeta', list: 'bullet', fw: 0.9, sizeFrac: 0.02, autoFit: false });
    assert.equal(drawn.filter((t) => t === '•').length, 2);
});

test('a checklist draws its boxes rather than writing a glyph', () => {
    // `✓` is not in every font stack, and a fallback would draw something else
    // on one of the two hosts — the same reason icons are paths.
    const plan = compile({ text: 'one\ntwo', list: 'check', fw: 0.9, sizeFrac: 0.02, autoFit: false });
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const written = ctx.calls.filter((c) => c.name === 'fillText').map((c) => c.args[0]);
    assert.deepEqual(written, ['one', 'two'], 'no marker glyph was written');
    assert.ok(ctx.calls.some((c) => c.name === 'stroke'), 'the boxes were stroked');
});

test('text is indented past the marker, and a paragraph is not', () => {
    const listX = compile({ text: 'alpha', list: 'bullet', fw: 0.8, sizeFrac: 0.02, autoFit: false });
    const plainX = compile({ text: 'alpha', list: 'none', fw: 0.8, sizeFrac: 0.02, autoFit: false });
    const at = (plan) => {
        const ctx = stubCtx();
        paintFrame(ctx, plan, 0);
        const body = ctx.calls.filter((c) => c.name === 'fillText' && c.args[0] === 'alpha')[0];
        return body.args[1];
    };
    assert.ok(at(listX) > at(plainX), 'the bulleted line starts further in');
});

// ── auto-fit ────────────────────────────────────────────────────────────────

test('auto-fit shrinks text that overflows its box', () => {
    const big = compile({ text: 'w '.repeat(200).trim(), fw: 0.5, fh: 0.1, sizeFrac: 0.05, autoFit: true });
    const ctx = stubCtx();
    paintFrame(ctx, big, 0);
    const drawn = ctx.calls.find((c) => c.name === 'fillText');
    const px = Number(/(\d+)px/.exec(drawn.font)?.[1]);
    assert.ok(px < 50, `stayed at ${px}px instead of shrinking`);
});

test('auto-fit never grows text that already fits', () => {
    // A block that grew to fill its box would change size as somebody typed,
    // which is the opposite of what a layout is for.
    const plan = compile({ text: 'short', fw: 0.8, fh: 0.8, sizeFrac: 0.03, autoFit: true });
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const px = Number(/(\d+)px/.exec(ctx.calls.find((c) => c.name === 'fillText').font)?.[1]);
    assert.equal(px, 30, 'the requested size is an upper bound, not a target');
});

test('auto-fit off leaves the size alone even when it overflows', () => {
    const plan = compile({ text: 'w '.repeat(200).trim(), fw: 0.5, fh: 0.1, sizeFrac: 0.05, autoFit: false });
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const px = Number(/(\d+)px/.exec(ctx.calls.find((c) => c.name === 'fillText').font)?.[1]);
    assert.equal(px, 50);
});

// ── tracking ────────────────────────────────────────────────────────────────

test('with no tracking a line is drawn in one call', () => {
    const drawn = drawnLines({ text: 'abcd', fw: 0.9, sizeFrac: 0.02, autoFit: false, tracking: 0 });
    assert.deepEqual(drawn, ['abcd']);
});

test('with tracking a line is drawn per character, not through letterSpacing', () => {
    // `ctx.letterSpacing` is recent and unevenly supported, so a host without
    // it would lay the same text out differently — and this renderer's claim is
    // that two hosts produce the same frames.
    const drawn = drawnLines({ text: 'abcd', fw: 0.9, sizeFrac: 0.02, autoFit: false, tracking: 0.1 });
    assert.deepEqual(drawn, ['a', 'b', 'c', 'd']);
});

// ── the layer, like any other ───────────────────────────────────────────────

test('a block reports bounds, so the editor can select and drag it', () => {
    const plan = compile({ text: 'alpha\nbeta', fw: 0.5, fx: 0.5, fy: 0.5, sizeFrac: 0.02, autoFit: false });
    const box = layerBounds(stubCtx(), plan, plan.layers[0]);
    assert.equal(box.w, 500, 'as wide as its fractional width');
    assert.ok(box.h > 0);
    assert.equal(box.x, 250, 'centred, like every other layer');
});

test('an explicit height wins over the measured one', () => {
    const plan = compile({ text: 'alpha', fw: 0.5, fh: 0.4, sizeFrac: 0.02, autoFit: false });
    const box = layerBounds(stubCtx(), plan, plan.layers[0]);
    assert.equal(box.h, 400);
});

test('a block with an unresolved token does not compile', () => {
    // The same rule word art and the text layer follow: a line with nothing
    // behind its token is not drawn empty, it is not drawn.
    const plan = compile({ text: '{nosuchtoken}' });
    assert.equal(plan.layers.length, 0);
});

test('the layout is memoised, and re-laid out when the text changes', () => {
    const plan = compile({ text: 'alpha', fw: 0.5, sizeFrac: 0.02, autoFit: false });
    const layer = plan.layers[0];
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    const first = layer._tb;
    paintFrame(ctx, plan, 1);
    assert.equal(layer._tb.value, first.value, 'a second frame reused the layout');
    layer.text = 'something else';
    paintFrame(ctx, plan, 2);
    assert.notEqual(layer._tb.value, first.value, 'a changed text re-laid out');
});
