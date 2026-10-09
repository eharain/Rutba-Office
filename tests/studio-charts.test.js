/**
 * Chart layers — the scale, the data shapes, and what each kind draws.
 *
 * The drawing itself is rectangles and arcs; what is worth testing is the
 * arithmetic around it, because that is what is wrong in a chart that looks
 * plausible. A bar scaled against the wrong maximum, a pie whose slices do not
 * close the circle, a nice-max that reads "47" — none of those look broken.
 *
 *   node --test tests/charts.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { applyLayerPatches, paintFrame, layerBounds, registeredLayerTypes } from '../packages/studio/src/index.js';
import { niceMax, readData, CHART_KINDS, CHART_LABELS, CHART_PALETTE, CHART_PALETTES } from '../packages/studio/src/charts.js';

function stubCtx() {
    const calls = [];
    const state = {};
    const record = (name) => (...args) => { calls.push({ name, args, fillStyle: state.fillStyle, globalCompositeOperation: state.globalCompositeOperation }); };
    const ctx = {
        calls,
        save: record('save'), restore: record('restore'),
        translate: record('translate'), rotate: record('rotate'), scale: record('scale'),
        beginPath: record('beginPath'), moveTo: record('moveTo'), lineTo: record('lineTo'),
        quadraticCurveTo: record('quadraticCurveTo'), bezierCurveTo: record('bezierCurveTo'),
        closePath: record('closePath'), ellipse: record('ellipse'), arc: record('arc'),
        rect: record('rect'), roundRect: record('roundRect'), clip: record('clip'),
        fill: record('fill'), stroke: record('stroke'), fillRect: record('fillRect'),
        setLineDash: record('setLineDash'), fillText: record('fillText'), strokeText: record('strokeText'),
        drawImage: record('drawImage'), putImageData: record('putImageData'),
        measureText: (s) => ({ width: String(s).length * 10 }),
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

const stubPlan = () => ({
    W: 1000, H: 1000, duration: 5, fps: 30,
    theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
    layers: [], images: [], videos: {}, assets: {}, context: {},
});

const compile = (patch) => {
    const plan = stubPlan();
    applyLayerPatches(plan, [{ id: 'c', type: 'chart', ...patch }]);
    return plan;
};

function drawn(patch) {
    const plan = compile(patch);
    const ctx = stubCtx();
    paintFrame(ctx, plan, 0);
    return ctx;
}

// ── the scale ───────────────────────────────────────────────────────────────

test('niceMax rounds up to a number a person would have chosen', () => {
    // The top gridline reading "47" is what makes a chart look measured rather
    // than designed.
    assert.equal(niceMax(47), 50);
    assert.equal(niceMax(230), 250);
    assert.equal(niceMax(1.7), 2);
    assert.equal(niceMax(9), 10);
    assert.equal(niceMax(10), 10);
    assert.equal(niceMax(0.04), 0.05);
});

test('niceMax never returns zero, so nothing divides by it', () => {
    assert.ok(niceMax(0) > 0);
    assert.ok(niceMax(-5) > 0);
    assert.ok(niceMax(NaN) > 0);
});

// ── data shapes ─────────────────────────────────────────────────────────────

test('a bare array of values is accepted as one series', () => {
    // Refusing the shorthand would mean the simplest possible chart needs the
    // most typing.
    const { series, count } = readData({ values: [1, 2, 3] });
    assert.equal(series.length, 1);
    assert.deepEqual(series[0].values, [1, 2, 3]);
    assert.equal(count, 3);
});

test('the full series form carries names and colours', () => {
    const { series } = readData({ series: [{ name: 'A', values: [1], color: '#fff' }] });
    assert.equal(series[0].name, 'A');
    assert.equal(series[0].color, '#fff');
});

test('missing labels are padded, not fatal', () => {
    // Labels and values are edited separately and drift; the chart draws what
    // it can rather than discarding everything over one missing label.
    const { labels } = readData({ values: [1, 2, 3], labels: ['a'] });
    assert.deepEqual(labels, ['a', '', '']);
});

test('non-numeric values become zero rather than NaN', () => {
    const { series } = readData({ values: [1, 'x', null, 4] });
    assert.deepEqual(series[0].values, [1, 0, 0, 4]);
});

test('an empty chart draws nothing rather than throwing', () => {
    const ctx = drawn({ values: [] });
    assert.ok(!ctx.calls.some((c) => c.name === 'fill'), 'nothing was painted');
});

// ── what each kind draws ────────────────────────────────────────────────────

test('every advertised kind paints without throwing', () => {
    for (const chart of CHART_KINDS) {
        const ctx = drawn({ chart, values: [3, 7, 5], labels: ['a', 'b', 'c'], fw: 0.8, fh: 0.5 });
        assert.ok(ctx.calls.length > 0, `${chart} drew nothing`);
    }
});

test('every kind has a label', () => {
    for (const k of CHART_KINDS) assert.ok(CHART_LABELS[k], `${k} has no label`);
});

test('columns scale against the nice maximum, not the raw one', () => {
    // Two charts with the same data and a different explicit max must draw the
    // same bar at different heights — which is what proves the scale is used.
    const a = drawn({ chart: 'column', values: [5], fw: 0.8, fh: 0.5, showGrid: false, showLabels: false, showValues: false });
    const b = drawn({ chart: 'column', values: [5], max: 100, fw: 0.8, fh: 0.5, showGrid: false, showLabels: false, showValues: false });
    const heightOf = (ctx) => {
        const move = ctx.calls.find((c) => c.name === 'moveTo');
        const line = ctx.calls.filter((c) => c.name === 'lineTo');
        return move.args[1] - line[0].args[1];
    };
    assert.ok(heightOf(a) > heightOf(b), 'a bigger axis makes the same value a shorter bar');
});

test('a pie closes the circle exactly', () => {
    // Slices that do not add to a full turn leave a wedge of background
    // showing, which reads as a missing category.
    const ctx = drawn({ chart: 'pie', values: [1, 2, 3], labels: ['a', 'b', 'c'], fw: 0.6, fh: 0.6 });
    const arcs = ctx.calls.filter((c) => c.name === 'arc');
    assert.equal(arcs.length, 3);
    const total = arcs.reduce((sum, a) => sum + (a.args[4] - a.args[3]), 0);
    assert.ok(Math.abs(total - Math.PI * 2) < 1e-9, `slices swept ${total}`);
});

test('a pie starts at twelve o’clock', () => {
    const ctx = drawn({ chart: 'pie', values: [1, 1], fw: 0.6, fh: 0.6 });
    const first = ctx.calls.find((c) => c.name === 'arc');
    assert.ok(Math.abs(first.args[3] + Math.PI / 2) < 1e-9, 'the first slice does not start at the top');
});

test('a donut punches its hole rather than painting over it', () => {
    // Painting a background-coloured disc would show as a solid circle the
    // moment the creative sits on a picture.
    const ctx = drawn({ chart: 'donut', values: [1, 2], fw: 0.6, fh: 0.6 });
    assert.ok(
        ctx.calls.some((c) => c.globalCompositeOperation === 'destination-out'),
        'the hole was filled, not cut',
    );
});

test('a line needs two points before it draws one', () => {
    const ctx = drawn({ chart: 'line', values: [5], fw: 0.8, fh: 0.5 });
    assert.ok(!ctx.calls.some((c) => c.name === 'stroke'), 'a single point is not a line');
});

test('a line chart draws every series', () => {
    const ctx = drawn({
        chart: 'line', fw: 0.8, fh: 0.5, showGrid: false,
        series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }],
    });
    assert.ok(ctx.calls.filter((c) => c.name === 'stroke').length >= 2);
});

test('an area chart fills translucently, so overlapping series stay readable', () => {
    // An opaque area chart with two series hides one of them completely.
    const plan = compile({ chart: 'area', values: [1, 2, 3], fw: 0.8, fh: 0.5 });
    assert.ok(plan.layers[0].fillOpacity > 0 && plan.layers[0].fillOpacity < 1);
});

test('values are shown on bars by default and hidden on lines', () => {
    // The number IS the point of a bar; a line with a number over every point
    // is unreadable.
    const bars = drawn({ chart: 'column', values: [3, 7], labels: ['a', 'b'], fw: 0.8, fh: 0.5 });
    const line = drawn({ chart: 'line', values: [3, 7], labels: ['a', 'b'], fw: 0.8, fh: 0.5 });
    const texts = (ctx) => ctx.calls.filter((c) => c.name === 'fillText').map((c) => c.args[0]);
    assert.ok(texts(bars).includes('3'), 'a bar shows its value');
    assert.ok(!texts(line).includes('3'), 'a line does not');
});

test('a prefix and suffix reach the drawn value', () => {
    const ctx = drawn({ chart: 'column', values: [42], prefix: '£', suffix: 'k', fw: 0.8, fh: 0.5 });
    assert.ok(ctx.calls.some((c) => c.name === 'fillText' && c.args[0] === '£42k'));
});

// ── the layer, like any other ───────────────────────────────────────────────

test('the chart type registers itself on import', () => {
    assert.ok(registeredLayerTypes().includes('chart'));
});

test('a chart reports bounds, so the editor can select and drag it', () => {
    const plan = compile({ values: [1], fw: 0.5, fh: 0.25, fx: 0.5, fy: 0.5 });
    const box = layerBounds(stubCtx(), plan, plan.layers[0]);
    assert.equal(box.w, 500);
    assert.equal(box.h, 250);
    assert.equal(box.x, 250, 'centred, like every other layer');
});

test('with no height a chart takes a sensible one from its width', () => {
    const plan = compile({ values: [1], fw: 0.5 });
    const box = layerBounds(stubCtx(), plan, plan.layers[0]);
    assert.ok(box.h > 0 && box.h < box.w, 'landscape by default');
});

test('the default palette has enough distinct colours to be told apart', () => {
    assert.ok(CHART_PALETTE.length >= 6);
    assert.equal(new Set(CHART_PALETTE).size, CHART_PALETTE.length, 'a repeated colour is two categories that look like one');
});

// ── one palette across the suite (M9) ───────────────────────────────────────

/**
 * The commons rule: the palette is the one thing every chart surface MUST
 * share, so a chart made here matches one made in Sheets. charts.js VENDORS
 * `@rutba/drawing`'s validated slots because this package is zero-dep by
 * contract — and this test is what makes the vendoring safe: the copy and
 * the source cannot drift without failing here.
 */
import { CATEGORICAL } from '@rutba/drawing/palette';

test('the vendored palette IS the suite palette, slot for slot, both variants', () => {
    assert.deepEqual(CHART_PALETTES.light, CATEGORICAL.light);
    assert.deepEqual(CHART_PALETTES.dark, CATEGORICAL.dark);
    assert.deepEqual(CHART_PALETTE, CATEGORICAL.light, 'the canonical export is the light set');
});

test('the theme decides which validated variant paints; an explicit palette still wins', () => {
    // Each variant is tuned for its surface — the dark set clears every
    // contrast check, the light set is the one Sheets draws on white.
    const dark = compile({ chart: 'column', values: [5] });
    dark.opts = { theme: 'dark' };
    const dctx = stubCtx();
    paintFrame(dctx, dark, 0);
    assert.ok(dctx.calls.some((c) => c.fillStyle === CHART_PALETTES.dark[0]), 'dark theme paints the dark slot');
    assert.ok(!dctx.calls.some((c) => c.fillStyle === CHART_PALETTES.light[0]), 'and not the light one');

    const pinned = compile({ chart: 'column', values: [5], palette: ['#123456'] });
    pinned.opts = { theme: 'dark' };
    const pctx = stubCtx();
    paintFrame(pctx, pinned, 0);
    assert.ok(pctx.calls.some((c) => c.fillStyle === '#123456'), 'a chosen palette is somebody having decided');
});

test('categories is accepted as an alias of labels — the word Sheets uses', () => {
    // One bound dataset must feed a sheet chart and a creative chart
    // unchanged (the T4 refinement already names this shape).
    const { labels } = readData({ categories: ['Q1', 'Q2'], values: [1, 2] });
    assert.deepEqual(labels, ['Q1', 'Q2']);
    const both = readData({ labels: ['a'], categories: ['b'], values: [1] });
    assert.deepEqual(both.labels, ['a'], 'labels wins — it is what the inspector writes');
});

test('an unknown chart kind falls back rather than failing to compile', () => {
    // The same forward-compatibility posture every other layer type keeps.
    const plan = compile({ chart: 'sunburst', values: [1] });
    assert.ok(CHART_KINDS.includes(plan.layers[0].chart));
});

// ── the editor's text data form ─────────────────────────────────────────────

/**
 * `label, value` per line is what the inspector edits.
 *
 * A grid of inputs is what a spreadsheet does; these charts are four to eight
 * points and typing "Mon, 12" beats tabbing between cells. What has to hold is
 * the ROUND TRIP — text in, arrays stored, text back — because the textarea is
 * rebuilt from the arrays and nothing keeps a second copy.
 */
import { dataToText, textToData } from '../packages/studio/src/chart-data.js';

test('text in, arrays out, text back — unchanged', () => {
    const text = 'Mon, 12\nTue, 19\nWed, 9';
    const { labels, values } = textToData(text);
    assert.deepEqual(labels, ['Mon', 'Tue', 'Wed']);
    assert.deepEqual(values, [12, 19, 9]);
    assert.equal(dataToText({ labels, values }), text);
});

test('a label containing a comma survives', () => {
    // The LAST comma splits. Splitting on the first would silently drop half of
    // any label with a comma in it — "Jan, Feb" is a real label.
    const { labels, values } = textToData('Jan, Feb, 12');
    assert.deepEqual(labels, ['Jan, Feb']);
    assert.deepEqual(values, [12]);
});

test('a pasted currency or percentage reads as its number', () => {
    // Copying a column out of a spreadsheet is the point of a text field; a
    // strict parse would turn every one of these into zero.
    const { values } = textToData('A, £1,200\nB, 24%\nC, 3.5');
    assert.deepEqual(values, [1200, 24, 3.5]);
});

test('values with no labels are accepted', () => {
    const { labels, values } = textToData('4\n8\n15');
    assert.deepEqual(values, [4, 8, 15]);
    assert.deepEqual(labels, ['', '', '']);
    assert.equal(dataToText({ labels, values }), '4\n8\n15');
});

test('blank lines are skipped rather than becoming zeroes', () => {
    const { values } = textToData('A, 1\n\n\nB, 2');
    assert.deepEqual(values, [1, 2]);
});

test('unparseable text becomes zero rather than NaN', () => {
    // NaN would propagate into the scale and blank the whole chart.
    const { values } = textToData('A, nonsense');
    assert.deepEqual(values, [0]);
});

test('the text form reads the series shape too', () => {
    // So switching a bound chart back to hand-typed data shows what it had.
    assert.equal(dataToText({ labels: ['A'], series: [{ values: [7] }] }), 'A, 7');
});

test('a negative value keeps its sign', () => {
    const { values } = textToData('Loss, -12');
    assert.deepEqual(values, [-12]);
});

test('a label with no number keeps its category rather than vanishing', () => {
    // A typo in one value should not silently delete the row it belongs to.
    const { labels, values } = textToData('Mon, 12\nTue, oops');
    assert.deepEqual(labels, ['Mon', 'Tue']);
    assert.deepEqual(values, [12, 0]);
});
